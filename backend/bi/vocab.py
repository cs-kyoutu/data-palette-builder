"""skills/bi/*.yaml 語彙のロードと解決ヘルパ。sql_builder / report_engine / prompts で共有。"""
from pathlib import Path

import yaml

_DIR = Path(__file__).parent.parent.parent / "skills" / "bi"


def _load(name: str):
    with open(_DIR / name, encoding="utf-8") as f:
        return yaml.safe_load(f)


AGG = _load("aggregation_methods.yaml")["methods"]
FILTERS = _load("filter_conditions.yaml")          # {string,numeric,date,multi_value}
PERIOD = _load("period_settings.yaml")
REPORTS = _load("report_types.yaml")
CHARTS = _load("chart_types.yaml")["charts"]

# 全フィルタ条件をフラットに (op解決用)
ALL_FILTERS = [e for group in FILTERS.values() for e in group]


def resolve(items: list[dict], token) -> dict | None:
    """key または ui の完全一致で語彙エントリを引く。"""
    if token is None:
        return None
    for e in items:
        if e.get("key") == token or e.get("ui") == token:
            return e
    return None


def ui_of(items: list[dict], token, default=None) -> str:
    """token(key/ui) に対応する ui 表示名。見つからなければ token そのまま。"""
    e = resolve(items, token)
    return e["ui"] if e else (default if default is not None else str(token))
