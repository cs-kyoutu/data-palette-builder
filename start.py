"""データパレット構築手順書ジェネレータ - 起動スクリプト"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("backend.app:app", host="0.0.0.0", port=8002, reload=True)
