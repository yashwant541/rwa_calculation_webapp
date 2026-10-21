import contextlib
import io
import json
import os
import sys
import types

import pandas as pd
import pytest

from webapp_core import split

ROOT = os.path.join(os.path.dirname(__file__), "..")


def _frame():
    return pd.DataFrame({
        "ID": [1, 2, 3, 4, 5],
        "TRANCHE_COLLATERAL_TYPE": ["Collaterized", " NOT  COLLATERIZED", "COLLATERIZED", None, "NOT COLLATERIZED"],
        "TRANCHE_GUARANTEE_TYPE": ["GUARANTEED", None, "guaranteed ", "NONE", "GUARANTEED"],
    })


def test_default_rules_split_into_three_datasets():
    out = {r["name"]: list(f["ID"]) for r, f in split.split_frame(_frame(), split.DEFAULT_SPLIT_RULES)}
    assert out["COLLATERIZED"] == [1, 3]            # equality, not substring of NOT COLLATERIZED
    assert out["NOT COLLATERIZED"] == [2, 5]        # whitespace + case normalised
    assert out["GUARANTEED"] == [1, 3, 5]           # independent filter, overlaps allowed


def test_missing_column_and_bad_rules():
    with pytest.raises(ValueError, match="not found"):
        split.split_frame(_frame()[["ID"]], split.DEFAULT_SPLIT_RULES)
    with pytest.raises(ValueError, match="pick a column"):
        split.clean_rules([{"name": "X", "column": "", "value": "A"}])
    with pytest.raises(ValueError, match="value"):
        split.clean_rules([{"name": "X", "column": "C", "value": " "}])
    with pytest.raises(ValueError, match="unique"):
        split.clean_rules([{"name": "A b", "column": "C", "value": "1"},
                           {"name": "a-b", "column": "C", "value": "2"}])


def test_file_name_and_workbook():
    rule = {"name": "NOT COLLATERIZED", "column": "C", "value": "x"}
    assert split.split_file_name("june exposures.xlsx", rule) == "june_exposures__NOT_COLLATERIZED.xlsx"
    back = pd.read_excel(io.BytesIO(split.frame_to_xlsx_bytes(_frame(), "NOT COLLATERIZED")), sheet_name=None)
    assert list(back) == ["NOT_COLLATERIZED"] and len(back["NOT_COLLATERIZED"]) == 5


# ---------------------------------------------------------------- backend glue
class _Stream(contextlib.AbstractContextManager):
    def __init__(self, data):
        self.data = data

    def read(self):
        return self.data

    def __exit__(self, *a):
        return False


STORE = {}


class FakeFolder:
    def __init__(self, folder_id):
        self.files = STORE.setdefault(folder_id, {})

    def list_paths_in_partition(self):
        return list(self.files)

    def get_download_stream(self, path):
        return _Stream(self.files[path])

    def upload_stream(self, path, stream):
        self.files[path] = stream.read()

    def delete_path(self, path):
        self.files.pop(path, None)

    def get_path_details(self, path):
        return {"size": len(self.files[path]), "lastModified": 1}


class FakeRecipe:
    ran = 0

    def run(self):
        FakeRecipe.ran += 1


def _xlsx(frame, sheet="Sheet1"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        frame.to_excel(w, sheet_name=sheet, index=False)
    return buf.getvalue()


@pytest.fixture
def client():
    from flask import Flask

    STORE.clear()
    FakeRecipe.ran = 0
    fake = types.ModuleType("dataiku")
    fake.Folder = FakeFolder
    fake.get_custom_variables = lambda: {}
    fake.api_client = lambda: types.SimpleNamespace(
        get_default_project=lambda: types.SimpleNamespace(get_recipe=lambda rid: FakeRecipe()))
    old = sys.modules.get("dataiku")
    sys.modules["dataiku"] = fake

    flask_app = Flask("t")
    ns = {"__name__": "rwa_backend_under_test", "app": flask_app}
    path = os.path.join(ROOT, "backend", "backend.py")
    exec(compile(open(path, encoding="utf-8").read(), path, "exec"), ns)

    cfg = ns["CATEGORY_CONFIG"]["corporate"]
    data, mapping = FakeFolder(cfg["data_folder_id"]), FakeFolder(cfg["mapping_folder_id"])
    f1 = _frame()
    f2 = _frame().assign(ID=[11, 12, 13, 14, 15])
    data.files["/june.xlsx"] = _xlsx(f1)
    data.files["/july.xlsx"] = _xlsx(f2)
    mapping.files["/map.xlsx"] = _xlsx(pd.DataFrame({"K": ["a"], "V": [1]}), sheet="Mapping")
    flask_app.config["_ns"] = ns
    yield flask_app.test_client(), cfg
    if old is not None:
        sys.modules["dataiku"] = old
    else:
        sys.modules.pop("dataiku", None)


def _post(c, url, body):
    r = c.post(url, data=json.dumps(body), content_type="application/json")
    return r.status_code, r.get_json()


def test_categories_expose_split_defaults(client):
    c, _ = client
    cat = next(x for x in c.get("/api/categories").get_json()["data"]["categories"] if x["key"] == "corporate")
    assert cat["split_enabled"] is False
    assert [r["column"] for r in cat["split_rules"]] == [
        "TRANCHE_COLLATERAL_TYPE", "TRANCHE_COLLATERAL_TYPE", "TRANCHE_GUARANTEE_TYPE"]


def test_preview_lists_columns_and_counts(client):
    c, _ = client
    code, body = _post(c, "/api/split-preview", {"category": "corporate", "rules": split.DEFAULT_SPLIT_RULES})
    d = body["data"]
    assert code == 200 and "TRANCHE_GUARANTEE_TYPE" in d["columns"]
    june = next(f for f in d["files"] if f["filename"] == "june.xlsx")
    assert [(x["name"], x["rows"]) for x in june["datasets"]] == [
        ("COLLATERIZED", 2), ("NOT COLLATERIZED", 2), ("GUARANTEED", 3)]

    bad = [{"name": "X", "column": "NOPE", "value": "1"}]
    d = _post(c, "/api/split-preview", {"category": "corporate", "rules": bad})[1]["data"]
    assert d["missing_columns"] == ["NOPE"]


def test_validate_flags_missing_split_column(client):
    c, _ = client
    rules = [{"name": "X", "column": "NOPE", "value": "1"}]
    d = _post(c, "/api/validate-inputs", {"category": "corporate", "selected_paths": ["/june.xlsx"],
                                          "split": {"enabled": True, "rules": rules}})[1]["data"]
    assert not d["valid"] and any("NOPE" in e for e in d["errors"])
    ok = _post(c, "/api/validate-inputs", {"category": "corporate", "selected_paths": ["/june.xlsx"],
                                           "split": {"enabled": True, "rules": split.DEFAULT_SPLIT_RULES}})[1]["data"]
    assert ok["valid"], ok["errors"]


def test_run_with_split_writes_six_files_and_points_manifest_at_them(client):
    c, cfg = client
    code, body = _post(c, "/api/run", {
        "category": "corporate", "selected_paths": ["/june.xlsx", "/july.xlsx"],
        "split": {"enabled": True, "rules": split.DEFAULT_SPLIT_RULES}})
    assert code == 200, body
    d = body["data"]
    assert len(d["split"]["datasets"]) == 6
    manifest = d["selection_manifest"]
    assert len(manifest["selected_paths"]) == 6
    assert manifest["source_paths"] == ["/june.xlsx", "/july.xlsx"]
    assert {x["split_name"] for x in manifest["split"]["datasets"]} == {
        "COLLATERIZED", "NOT COLLATERIZED", "GUARANTEED"}
    files = FakeFolder(cfg["data_folder_id"]).files
    one = pd.read_excel(io.BytesIO(files["/_webapp_metadata/split/june__COLLATERIZED.xlsx"]))
    assert list(one["ID"]) == [1, 3]
    assert FakeRecipe.ran == 1

    # a second run replaces the previous split files instead of piling them up
    _post(c, "/api/run", {"category": "corporate", "selected_paths": ["/june.xlsx"],
                          "split": {"enabled": True, "rules": split.DEFAULT_SPLIT_RULES}})
    left = [p for p in files if p.startswith("/_webapp_metadata/split/")]
    assert len(left) == 3


def test_edited_rule_changes_the_dataset(client):
    c, _ = client
    rules = [{"name": "NO GUARANTEE", "column": "TRANCHE_GUARANTEE_TYPE", "value": "none"}]
    d = _post(c, "/api/run", {"category": "corporate", "selected_paths": ["/june.xlsx"],
                              "split": {"enabled": True, "rules": rules}})[1]["data"]
    assert [(x["split_name"], x["row_count"]) for x in d["split"]["datasets"]] == [("NO GUARANTEE", 1)]


def test_run_without_split_is_unchanged(client):
    c, cfg = client
    d = _post(c, "/api/run", {"category": "corporate", "selected_paths": ["/june.xlsx"]})[1]["data"]
    assert d["selection_manifest"]["selected_paths"] == ["/june.xlsx"]
    assert "split" not in d["selection_manifest"] and d["split"] is None
    assert not [p for p in FakeFolder(cfg["data_folder_id"]).files if "/split/" in p]
