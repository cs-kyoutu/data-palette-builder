import os
import uvicorn

if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    uvicorn.run("backend.app_phase2:app", host="0.0.0.0", port=8004, reload=True)
