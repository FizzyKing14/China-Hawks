"""
Fizz AI — 你的私人 AI 助手，由本地开源模型驱动

快速启动:
    # 1. 安装依赖
    pip install -r fizz/requirements.txt

    # 2. 安装 Ollama（首次）: https://ollama.com/download
    #    然后下载模型:
    ollama pull qwen2.5          # 推荐中文模型
    # 或者:
    ollama pull llama3.2

    # 3. 启动 Fizz
    uvicorn fizz.main:app --reload --port 7860

    # 在浏览器打开 http://localhost:7860
"""

from __future__ import annotations

import json
from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel

try:
    import ollama as _ollama
    _OLLAMA_OK = True
except ImportError:
    _OLLAMA_OK = False

app = FastAPI(title="Fizz AI", version="1.0.0")

_STATIC = Path(__file__).parent / "static"

_DEFAULT_SYSTEM = (
    "你是 Fizz，一个聪明、友好、乐于助人的 AI 助手。"
    "你完全运行在用户本地的电脑上，数据不会上传到任何服务器，绝对隐私安全。"
    "用户用什么语言提问，你就用什么语言回答。"
    "回答要准确、简洁清晰，代码和结构化内容可以使用 Markdown 格式。"
)


@app.get("/", response_class=HTMLResponse)
async def index():
    return (_STATIC / "index.html").read_text(encoding="utf-8")


@app.get("/api/health")
async def health():
    return {"name": "Fizz", "version": "1.0.0", "ollama_available": _OLLAMA_OK}


@app.get("/api/models")
async def list_models():
    if not _OLLAMA_OK:
        return {"models": [], "error": "ollama 未安装，请运行: pip install ollama"}
    try:
        data = _ollama.list()
        raw = data.get("models") or []
        models = [m.get("model") or m.get("name", "") for m in raw]
        return {"models": [m for m in models if m]}
    except Exception as exc:
        return {"models": [], "error": f"无法连接 Ollama 服务: {exc}"}


class Message(BaseModel):
    role: str      # "user" | "assistant"
    content: str


class ChatRequest(BaseModel):
    model: str = "qwen2.5"
    messages: list[Message]
    system: str = _DEFAULT_SYSTEM


@app.post("/api/chat")
async def chat(req: ChatRequest):
    if not _OLLAMA_OK:
        async def _err():
            yield f"data: {json.dumps({'error': 'ollama 未安装，请运行 pip install ollama'})}\n\n"
            yield "data: [DONE]\n\n"
        return StreamingResponse(_err(), media_type="text/event-stream")

    msgs = [{"role": "system", "content": req.system}]
    msgs += [{"role": m.role, "content": m.content} for m in req.messages]

    def _stream():
        try:
            for chunk in _ollama.chat(model=req.model, messages=msgs, stream=True):
                token = chunk.get("message", {}).get("content", "")
                if token:
                    yield f"data: {json.dumps({'content': token})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'error': str(exc)})}\n\n"
        yield "data: [DONE]\n\n"

    return StreamingResponse(_stream(), media_type="text/event-stream")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("fizz.main:app", host="0.0.0.0", port=7860, reload=True)
