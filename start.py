"""データパレット構築手順書ジェネレータ - 起動スクリプト"""
import os
from pathlib import Path
import uvicorn

if __name__ == "__main__":
    os.chdir(Path(__file__).parent)
    uvicorn.run("backend.app:app", host="0.0.0.0", port=8002)
