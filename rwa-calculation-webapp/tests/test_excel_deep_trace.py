import excel_deep_trace as edt


def test_parse_excel_shapes():
    node, _ = edt.parse_excel("=IF(ABS(E2-F2)<=0.01,\"MATCH\",\"MISMATCH\")", "RWA")
    assert isinstance(node, edt.If)
    assert isinstance(node.then, edt.Str) and node.then.value == "MATCH"
    v, _ = edt.parse_excel("=VLOOKUP(C2,Ratings!$A$2:$B$5,2,FALSE)", "RWA")
    assert isinstance(v, edt.Call) and v.name == "VLOOKUP"
    assert isinstance(v.args[1], edt.RangeRef) and v.args[1].sheet == "Ratings"


def test_scan_summary(rwa_engine):
    s = rwa_engine.scan()
    assert s["total_formula_cells"] == 9          # 3 rows x 3 formula columns
    assert "VLOOKUP" in s["function_usage"] and "IF" in s["function_usage"]
    assert s["lookup_cells"] == 3 and s["conditional_cells"] == 3
    assert s["families"]["lookup"] == 3


def test_normalize_key_and_value(rwa_engine):
    assert rwa_engine._normalize_key("rwa!e2") == "RWA!E2"
    assert rwa_engine._normalize_key("E2") == "RWA!E2"
    assert abs(rwa_engine.value("RWA!E2") - 160.0) < 1e-9
    assert rwa_engine.value("RWA!G2") == "MISMATCH"


def test_record_metadata(rwa_engine):
    b2 = rwa_engine.records["RWA!B2"]
    assert b2.is_lookup and b2.functions == ["VLOOKUP"] and b2.family == "lookup"
    e2 = rwa_engine.records["RWA!E2"]
    assert {"RWA!A2", "RWA!B2", "RWA!D2"} <= e2.precedents


def test_trace_precedents(rwa_engine):
    g = rwa_engine.trace("RWA!E2", "precedents", max_depth=5)
    ids = {n["id"] for n in g["nodes"]}
    assert {"RWA!E2", "RWA!A2", "RWA!B2", "RWA!D2"} <= ids
    # B2 -> depends on the Ratings range corners
    assert any(n["id"].startswith("Ratings!") for n in g["nodes"])
    assert ("RWA!B2", "RWA!E2") in {(e["from"], e["to"]) for e in g["edges"]}


def test_trace_dependents(rwa_engine):
    g = rwa_engine.trace("RWA!B2", "dependents", max_depth=5)
    assert "RWA!E2" in {n["id"] for n in g["nodes"]}


def test_branch_analysis(rwa_engine):
    r2 = rwa_engine.branch_analysis("RWA!G2")
    assert r2["branches"][0]["branch_taken"] == "FALSE"     # 160 vs 200 -> mismatch
    r3 = rwa_engine.branch_analysis("RWA!G3")
    assert r3["branches"][0]["branch_taken"] == "TRUE"      # 500 vs 500 -> match


def test_lookup_analysis(rwa_engine):
    la = rwa_engine.lookup_analysis("RWA!B2")
    d = la["lookups"][0]
    assert d["function"] == "VLOOKUP"
    assert d["match_found"] and d["matched_key"] == "AAA"
    assert abs(d["returned_value"] - 0.2) < 1e-9


def test_expression_tree_reproduces_excel(rwa_engine):
    t = rwa_engine.expression_tree("RWA!E2")
    assert t["is_formula"] and abs(t["value"] - 160.0) < 1e-9
    assert t["tracer_reproduced_excel"] is True
    # the VLOOKUP is reached by recursing into B2
    def has_vlookup(n):
        if n.get("kind") == "lookup":
            return True
        return any(has_vlookup(c) for c in n.get("children", []))
    assert has_vlookup(t["root"])


def test_expression_tree_branch_and_narrative(rwa_engine):
    t = rwa_engine.expression_tree("RWA!G2")
    assert t["root"]["kind"] == "if"
    assert t["root"]["detail"]["branch"] == "false"
    text = " ".join(t["narrative"])
    assert "MISMATCH" in text or "FALSE branch" in text
    assert t["value"] == "MISMATCH"


def test_row_journey(rwa_engine):
    j = rwa_engine.row_journey("RWA", 2)
    assert j["formula_cells"] == 3
    cells = {c["cell"] for c in j["cells"]}
    assert cells == {"B2", "E2", "G2"}
    inputs = {c["cell"]: c["value"] for c in j["inputs"]}
    assert inputs["A2"] == 1000 and inputs["C2"] == "AAA" and inputs["F2"] == 200
    g2 = next(c for c in j["cells"] if c["cell"] == "G2")
    assert g2["is_conditional"] and g2["branches"][0]["branch_taken"] == "FALSE"


def test_compare_cells(rwa_engine):
    c = rwa_engine.compare_cells("RWA!E2", "RWA!F2")
    assert c["values_match"] is False
    assert abs(c["difference"] - (-40.0)) < 1e-9
    assert "RWA!A2" in c["only_left_precedents"] or "RWA!A2" in c["shared_precedents"]


def test_findings_are_listed(rwa_engine):
    codes = {f.code for f in rwa_engine.findings}
    # 0.01 tolerance literal mixed with cell refs in the STATUS formula
    assert "HARDCODED_CONSTANT" in codes
    for f in rwa_engine.findings:
        assert f.severity in ("HIGH", "MEDIUM", "LOW", "INFO")


def test_input_cell_expression_tree(rwa_engine):
    t = rwa_engine.expression_tree("RWA!A2")
    assert t["is_formula"] is False and t["value"] == 1000


def test_sheet_columns_and_mismatch_rows(rwa_engine):
    info = rwa_engine.sheet_columns("RWA")
    assert info["header_row"] == 1
    names = [c["name"] for c in info["columns"]]
    assert "CALCULATED_RWA" in names and "SOURCE_RWA" in names

    mm = rwa_engine.mismatch_rows("RWA", "CALCULATED_RWA", "SOURCE_RWA", tolerance=0.01)
    # row 2 only: 160 vs 200 ; rows 3 & 4 match
    assert [r["row"] for r in mm["rows"]] == [2]
    assert mm["rows"][0]["diff"] == -40.0
    # also resolvable by column letter
    assert rwa_engine.mismatch_rows("RWA", "E", "F")["total"] == 1


def test_family_arithmetic_for_rounded_expression():
    node, _ = edt.parse_excel("=ROUND(A2*B2*(1-D2),2)", "S")
    assert edt._family({"ROUND"}, node) == "arithmetic"
    node2, _ = edt.parse_excel('=IF(A2>0,"y","n")', "S")
    assert edt._family({"IF"}, node2) == "conditional"
