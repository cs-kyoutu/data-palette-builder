"""b→dash BI モード — データパレット(加工)とは別ネームスペース。

横断インフラ(セッション/Claude/DB/認証)は backend/_shared.py を再利用し、
BI 固有の語彙(skills/bi/*.yaml)・SQL ビルダー・手順書レンダラ・Excel をここに隔離する。
初期スコープ = カスタムレポート + セグメント (BI_.md / docs/bi_mode_design.md)。
"""
