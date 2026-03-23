"""データパレット構築 設計書・手順書ジェネレーター - 起動スクリプト"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("backend.app_combined:app", host="0.0.0.0", port=8002, reload=True)
