"""依赖健康审计。

在已构建的单元格依赖图上扫描四类问题：

- orphan_formula（info）   ：公式单元格未被任何其他公式引用（孤立公式 / 疑似死计算）
- empty_reference（warning）：被公式引用的目标单元格为空（既无公式也无缓存值）
- missing_sheet（error）   ：跨表公式指向不存在的工作表
- out_of_sheet_range（warning）：跨表引用落在目标工作表已用范围之外
- high_fanout（warning）   ：扇出（直接引用的单元格数）超过阈值

每条问题给出：级别、类型、所在工作表、单元格位置、说明、证据边。
"""
from __future__ import annotations

from collections import defaultdict

from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

from .formula_parser import RawRef
from .graph import WorkbookGraph, display_id, split_node

# 级别顺序（用于排序与计数）
LEVELS = ("error", "warning", "info")

LEVEL_LABELS = {
    "error": "错误",
    "warning": "警告",
    "info": "提示",
}

TYPE_LABELS = {
    "orphan_formula": "孤立公式",
    "empty_reference": "引用空单元格",
    "missing_sheet": "引用不存在的工作表",
    "out_of_sheet_range": "跨表引用超出已用范围",
    "high_fanout": "扇出过大",
}

# 每条问题证据边列表的上限
MAX_EVIDENCE = 20
DEFAULT_FANOUT_THRESHOLD = 50
FANOUT_MIN = 1
FANOUT_MAX = 100000


def audit_graph(
    graph: WorkbookGraph,
    fanout_threshold: int = DEFAULT_FANOUT_THRESHOLD,
    sheet: str | None = None,
    level: str | None = None,
) -> dict:
    """对依赖图执行健康审计，返回含汇总计数的完整结果（不受筛选影响的全量统计）。"""
    bounds = {s.name: (s.max_row, s.max_col) for s in graph.sheets}
    issues: list[dict] = []

    # 1) 缺失工作表 / 跨表超出已用范围：直接检查公式中的原始引用
    for node, refs in graph.raw_refs.items():
        owner_sheet, coord = split_node(node)
        seen_texts: set[str] = set()
        for ref in refs:
            if ref.sheet is None:
                continue  # 仅审计跨表引用
            if ref.text in seen_texts:
                continue
            seen_texts.add(ref.text)
            if ref.sheet not in bounds:
                issues.append(_missing_sheet(node, owner_sheet, coord, ref))
                continue
            over = _out_of_range_count(ref, bounds[ref.sheet])
            if over:
                issues.append(_out_of_range(node, owner_sheet, coord, ref, bounds[ref.sheet], over))

    # 2) 被引用却为空的单元格：聚合所有指向它的公式
    empty_referrers: dict[str, list[str]] = defaultdict(list)
    for src, targets in graph.precedents.items():
        src_sheet = split_node(src)[0]
        for tnode in targets:
            if tnode in graph.formulas or tnode in graph.values:
                continue
            tsheet, tcoord = split_node(tnode)
            # 仅跨表且落在目标已用范围之外的引用已记为 out_of_sheet_range，不重复计数；
            # 同表越界引用（如本表 C2 尚无内容）仍按“引用空单元格”处理
            if tsheet != src_sheet:
                b = bounds.get(tsheet)
                if b is not None and _coord_out_of_bounds(tcoord, b):
                    continue
            empty_referrers[tnode].append(src)
    for tnode in sorted(empty_referrers):
        tsheet, tcoord = split_node(tnode)
        referrers = sorted(empty_referrers[tnode])
        issues.append({
            "level": "warning",
            "type": "empty_reference",
            "type_label": TYPE_LABELS["empty_reference"],
            "sheet": tsheet,
            "cell": tnode,
            "cell_display": display_id(tsheet, tcoord),
            "message": (
                f"单元格 {display_id(tsheet, tcoord)} 为空（无公式、无值），"
                f"却被 {len(referrers)} 个公式引用"
            ),
            "evidence": [
                {"from": display_id(*split_node(r)), "to": display_id(tsheet, tcoord)}
                for r in referrers[:MAX_EVIDENCE]
            ],
            "evidence_total": len(referrers),
        })

    # 3) 孤立公式：未被任何公式引用的公式单元格
    for node in sorted(graph.formulas):
        if graph.dependents.get(node):
            continue
        owner_sheet, coord = split_node(node)
        issues.append({
            "level": "info",
            "type": "orphan_formula",
            "type_label": TYPE_LABELS["orphan_formula"],
            "sheet": owner_sheet,
            "cell": node,
            "cell_display": display_id(owner_sheet, coord),
            "message": f"公式 {display_id(owner_sheet, coord)} 未被任何其他公式引用，疑似孤立/废弃计算",
            "evidence": [],
            "evidence_total": 0,
        })

    # 4) 扇出过大的公式节点
    for node in sorted(graph.formulas):
        targets = graph.precedents.get(node, set())
        fanout = len(targets)
        if fanout < fanout_threshold:
            continue
        owner_sheet, coord = split_node(node)
        sample = sorted(targets)[:MAX_EVIDENCE]
        issues.append({
            "level": "warning",
            "type": "high_fanout",
            "type_label": TYPE_LABELS["high_fanout"],
            "sheet": owner_sheet,
            "cell": node,
            "cell_display": display_id(owner_sheet, coord),
            "message": (
                f"公式 {display_id(owner_sheet, coord)} 直接引用 {fanout} 个单元格，"
                f"超过扇出阈值 {fanout_threshold}"
            ),
            "fanout": fanout,
            "threshold": fanout_threshold,
            "evidence": [
                {"from": display_id(owner_sheet, coord), "to": display_id(*split_node(t))}
                for t in sample
            ],
            "evidence_total": fanout,
        })

    issues.sort(key=_sort_key)
    counts = {"error": 0, "warning": 0, "info": 0}
    counts_by_type: dict[str, int] = {}
    for it in issues:
        counts[it["level"]] += 1
        counts_by_type[it["type"]] = counts_by_type.get(it["type"], 0) + 1

    filtered = [
        it for it in issues
        if (sheet is None or it["sheet"] == sheet)
        and (level is None or it["level"] == level)
    ]
    return {
        "fanout_threshold": fanout_threshold,
        "filters": {"sheet": sheet, "level": level},
        "total": len(filtered),
        "counts": {**counts, "all": len(issues)},
        "counts_by_type": counts_by_type,
        "issues": filtered,
    }


# ------------------------------------------------------------------ 辅助
def _issue(node: str, level: str, itype: str, message: str,
           evidence: list[dict], evidence_total: int | None = None) -> dict:
    owner_sheet, coord = split_node(node)
    return {
        "level": level,
        "type": itype,
        "type_label": TYPE_LABELS[itype],
        "sheet": owner_sheet,
        "cell": node,
        "cell_display": display_id(owner_sheet, coord),
        "message": message,
        "evidence": evidence[:MAX_EVIDENCE],
        "evidence_total": len(evidence) if evidence_total is None else evidence_total,
    }


def _missing_sheet(node: str, owner_sheet: str, coord: str, ref: RawRef) -> dict:
    target = f"{ref.sheet}!{ref.text.split('!', 1)[-1]}"
    return _issue(
        node,
        "error",
        "missing_sheet",
        f"公式中的跨表引用 {target} 指向不存在的工作表「{ref.sheet}」",
        [{"from": display_id(owner_sheet, coord), "to": target, "ref": ref.text,
          "missing_sheet": ref.sheet}],
    )


def _out_of_range(node: str, owner_sheet: str, coord: str, ref: RawRef,
                  bounds: tuple[int, int], over: int) -> dict:
    max_row, max_col = bounds
    return _issue(
        node,
        "warning",
        "out_of_sheet_range",
        (
            f"跨表引用 {ref.text} 超出工作表「{ref.sheet}」已用范围 "
            f"（{max_row} 行 × {max_col} 列），其中 {over} 个单元格在已用范围之外"
        ),
        [{"from": display_id(owner_sheet, coord), "to": ref.text, "ref": ref.text,
          "target_sheet": ref.sheet,
          "sheet_max_row": max_row, "sheet_max_col": max_col,
          "out_of_range_cells": over}],
    )


def _out_of_range_count(ref: RawRef, bounds: tuple[int, int]) -> int:
    """引用区域落在目标工作表已用范围之外的单元格数量；完全在范围内返回 0。"""
    max_row, max_col = bounds
    if ref.kind == "cell":
        return int(ref.row1 > max_row or ref.col1 > max_col)
    if ref.kind == "range":
        r1, r2, c1, c2 = ref.row1, ref.row2, ref.col1, ref.col2
    elif ref.kind == "cols":
        r1, r2, c1, c2 = 1, max_row, ref.col1, ref.col2
    else:  # rows
        r1, r2, c1, c2 = ref.row1, ref.row2, 1, max_col
    if r2 <= max_row and c2 <= max_col:
        return 0
    total = (r2 - r1 + 1) * (c2 - c1 + 1)
    ir = (min(r2, max_row) - r1 + 1)
    ic = (min(c2, max_col) - c1 + 1)
    inside = ir * ic if ir > 0 and ic > 0 else 0
    return total - inside


def _coord_out_of_bounds(coord: str, bounds: tuple[int, int]) -> bool:
    max_row, max_col = bounds
    col_letters, row = coordinate_from_string(coord)
    return row > max_row or column_index_from_string(col_letters) > max_col


def _sort_key(issue: dict) -> tuple:
    level_rank = {lvl: i for i, lvl in enumerate(LEVELS)}
    sheet, coord = split_node(issue["cell"])
    col_letters, row = coordinate_from_string(coord)
    return (
        level_rank[issue["level"]],
        issue["sheet"],
        row,
        column_index_from_string(col_letters),
        issue["type"],
    )
