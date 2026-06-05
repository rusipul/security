import json, os
from config import load_keywords, load_settings, save_settings

def test_load_keywords_returns_design_keywords():
    kw = load_keywords("keywords.json")
    assert "DB" in kw["design_keywords"]
    assert "design_extensions" in kw
    assert kw["zip_size_threshold_mb"] == 10

def test_settings_roundtrip(tmp_path):
    cfg = tmp_path / "config.json"
    save_settings({"api_key": "sk-test", "top_n": 15}, str(cfg))
    loaded = load_settings(str(cfg))
    assert loaded["api_key"] == "sk-test"
    assert loaded["top_n"] == 15

def test_load_settings_returns_defaults_when_missing(tmp_path):
    cfg = tmp_path / "no_config.json"
    loaded = load_settings(str(cfg))
    assert loaded["top_n"] == 10
    assert loaded["api_key"] == ""
