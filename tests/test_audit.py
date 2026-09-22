"""依赖健康审计的核心逻辑测试。运行：python -m pytest tests/ -v"""
import io

from openpyxl import Workbook, load_workbook

from app.audit import audit_graph
from app.graph import WorkbookGraph


def build_audit_graph():
    """构造覆盖全部审计问题类型的工作簿。

    S1（已用范围约 A1:D8）：
      A1=1, B1==A1*2                       B1 被 D1 引用，非孤立
      C1==A1+1                             孤立公式
      D1==B1+C8                            C8 同表越界空引用；D1 扇出 2
      D2==Ghost!A1                         指向不存在的工作表
      D3==S2!E2                            跨表超已用范围（S2 仅到 C3）
      D4==SUM(S2!A1:C2)                    扇出 6（阈值 5 时 high_fanout）
      D5==S2!B2+S2!B3                      引用范围内空格 B3（B2 有值）
    S2：A1=1 B1=2 A2=3 B2=4 C2=5（B3 为空但在已用范围内）
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["B1"] = "=A1*2"
    ws["C1"] = "=A1+1"
    ws["D1"] = "=B1+C8"
    ws["D2"] = "=Ghost!A1"
    ws["D3"] = "=S2!E2"
    ws["D4"] = "=SUM(S2!A1:C2)"
    ws["D5"] = "=S2!B2+S2!B3"
    ws2 = wb.create_sheet("S2")
    ws2["A1"] = 1
    ws2["B1"] = 2
    ws2["A2"] = 3
    ws2["B2"] = 4
    ws2["C2"] = 5
    ws2["A3"] = "预留行"  # 使 S2 已用范围覆盖第 3 行，B3 成为“范围内空格”
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv)


def issue_map(result, itype):
    return {i["cell"]: i for i in result["issues"] if i["type"] == itype}


def test_all_issue_types_detected():
    g = build_audit_graph()
    res = audit_graph(g, fanout_threshold=5)
    by_type = res["counts_by_type"]
    assert by_type.get("missing_sheet") >= 1
    assert by_type.get("out_of_sheet_range") >= 1
    assert by_type.get("high_fanout") >= 1
    assert by_type.get("empty_reference") >= 1
    assert by_type.get("orphan_formula") >= 1
    assert res["counts"]["all"] == len(res["issues"])


def test_missing_sheet_is_error_with_edge():
    g = build_audit_graph()
    issues = issue_map(audit_graph(g, fanout_threshold=5), "missing_sheet")
    it = issues["S1!D2"]
    assert it["level"] == "error"
    assert it["sheet"] == "S1"
    assert it["evidence"][0]["from"] == "S1!D2"
    assert "Ghost" in it["evidence"][0]["to"]
    assert it["evidence"][0]["missing_sheet"] == "Ghost"


def test_out_of_sheet_range_is_warning():
    g = build_audit_graph()
    issues = issue_map(audit_graph(g, fanout_threshold=5), "out_of_sheet_range")
    it = issues["S1!D3"]
    assert it["level"] == "warning"
    assert it["evidence"][0]["target_sheet"] == "S2"
    assert it["evidence"][0]["out_of_range_cells"] == 1


def test_out_of_range_count_for_ranges():
    g = build_audit_graph()
    res = audit_graph(g, fanout_threshold=100)  # 阈值拉高，只剩结构性问题
    oor = {i["cell"]: i for i in res["issues"] if i["type"] == "out_of_sheet_range"}
    assert "S1!D3" in oor  # 单格 E2：列超界
    # D4 区域 A1:C2 完全在 S2 已用范围内，不报
    assert "S1!D4" not in oor


def test_empty_references_in_and_out_of_range():
    g = build_audit_graph()
    issues = issue_map(audit_graph(g, fanout_threshold=5), "empty_reference")
    # 同表越界空格：仍算空引用
    assert "S1!C8" in issues
    assert issues["S1!C8"]["evidence"][0]["from"] == "S1!D1"
    # 跨表、范围内空格
    assert "S2!B3" in issues
    assert issues["S2!B3"]["evidence"][0]["from"] == "S1!D5"
    # 跨表越界（S2!E2）已计入 out_of_sheet_range，不再重复为空引用
    assert "S2!E2" not in issues


def test_orphan_formula_info():
    g = build_audit_graph()
    issues = issue_map(audit_graph(g, fanout_threshold=5), "orphan_formula")
    assert "S1!C1" in issues
    assert issues["S1!C1"]["level"] == "info"
    # B1 被 D1 引用，不是孤立公式
    assert "S1!B1" not in issues


def test_high_fanout_respects_threshold():
    g = build_audit_graph()
    res_low = audit_graph(g, fanout_threshold=5)
    issues = issue_map(res_low, "high_fanout")
    it = issues["S1!D4"]
    assert it["level"] == "warning"
    assert it["fanout"] == 6 and it["threshold"] == 5
    assert it["evidence_total"] == 6 and len(it["evidence"]) == 6
    # 提高阈值后不再报告
    res_high = audit_graph(g, fanout_threshold=50)
    assert "S1!D4" not in issue_map(res_high, "high_fanout")


def test_filters_level_and_sheet_keep_global_counts():
    g = build_audit_graph()
    full = audit_graph(g, fanout_threshold=5)
    warn = audit_graph(g, fanout_threshold=5, level="warning")
    assert all(i["level"] == "warning" for i in warn["issues"])
    assert warn["total"] == len(warn["issues"])
    assert warn["counts"] == full["counts"]  # 汇总计数不受筛选影响

    s2 = audit_graph(g, fanout_threshold=5, sheet="S2")
    assert all(i["sheet"] == "S2" for i in s2["issues"])
    # 归属 S2 的问题：空引用 S2!B3
    assert "S2!B3" in {i["cell"] for i in s2["issues"]}
    # 指向不存在工作表的公式归属 S1
    assert all(i["cell"] != "S1!D2" for i in s2["issues"])


def test_issues_sorted_by_level():
    g = build_audit_graph()
    levels = [i["level"] for i in audit_graph(g, fanout_threshold=5)["issues"]]
    rank = {"error": 0, "warning": 1, "info": 2}
    assert levels == sorted(levels, key=lambda lv: rank[lv])


def test_clean_workbook_has_no_issues():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["B1"] = "=A1+1"  # B1 也会被视为孤立公式
    ws2 = wb.create_sheet("S2")
    ws2["A1"] = "=S1!B1*2"  # B1 被引用 -> 无孤立；引用均合法 -> 零结构性问题
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    g = WorkbookGraph.from_workbooks(wbf, wbv)
    res = audit_graph(g, fanout_threshold=50)
    assert res["counts"]["error"] == 0
    assert res["counts"]["warning"] == 0
    # S2!A1 作为最终结果未被引用，属 info 级孤立公式；不应有更严重问题
    orphan_sheets = {i["sheet"] for i in res["issues"] if i["type"] == "orphan_formula"}
    assert orphan_sheets == {"S2"}
