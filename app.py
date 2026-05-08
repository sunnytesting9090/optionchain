from __future__ import annotations

import asyncio
import queue
import threading
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, JSONResponse

from update_simple_option_chain import (
    build_expiry_map,
    build_kite_client,
    build_runtime_plan,
    load_access_token,
    load_instrument_universe,
    log,
    seed_initial_ticks,
)
from zerodha_option_chain_server import (
    CLIENTS,
    broadcast_loop,
    client_handler,
    initial_messages,
    start_kite_stream,
    token_metadata,
)


BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_PATH = BASE_DIR / "option_chain.html"


class FastAPIWebSocketAdapter:
    def __init__(self, websocket: WebSocket) -> None:
        self.websocket = websocket

    async def send(self, message: str) -> None:
        await self.websocket.send_text(message)

    def __aiter__(self) -> "FastAPIWebSocketAdapter":
        return self

    async def __anext__(self) -> str:
        try:
            return await self.websocket.receive_text()
        except WebSocketDisconnect:
            raise StopAsyncIteration from None


def build_app_state() -> dict[str, Any]:
    access_token = load_access_token()
    kite = build_kite_client(access_token)
    instrument_frame = load_instrument_universe(kite)
    plan = build_runtime_plan(kite, instrument_frame)
    latest_ticks = seed_initial_ticks(kite, plan)
    metadata = token_metadata(plan)
    expiry_map = build_expiry_map(instrument_frame)
    initial_payload = initial_messages(latest_ticks, metadata)

    tick_queue: queue.Queue[list[dict[str, Any]]] = queue.Queue()
    stop_event = threading.Event()
    ticker_ws_holder: dict[str, Any] = {}
    ticker = start_kite_stream(access_token, plan, tick_queue, stop_event, ticker_ws_holder)
    context = {
        "kite": kite,
        "instrument_frame": instrument_frame,
        "metadata": metadata,
        "expiry_map": expiry_map,
        "ticker_ws_holder": ticker_ws_holder,
        "latest_expiry": {},
    }

    return {
        "kite": kite,
        "plan": plan,
        "latest_ticks": latest_ticks,
        "metadata": metadata,
        "initial_payload": initial_payload,
        "tick_queue": tick_queue,
        "stop_event": stop_event,
        "ticker": ticker,
        "context": context,
    }


@asynccontextmanager
async def lifespan(app: FastAPI):
    state = build_app_state()
    app.state.option_chain = state
    broadcaster = asyncio.create_task(
        broadcast_loop(
            state["kite"],
            state["tick_queue"],
            state["latest_ticks"],
            state["metadata"],
            state["stop_event"],
        )
    )
    app.state.option_chain_broadcaster = broadcaster
    log("FastAPI option-chain server ready.")
    try:
        yield
    finally:
        state["stop_event"].set()
        broadcaster.cancel()
        try:
            await broadcaster
        except asyncio.CancelledError:
            pass
        state["ticker"].close()
        CLIENTS.clear()


app = FastAPI(title="Zerodha Option Chain", lifespan=lifespan)


@app.get("/")
async def dashboard() -> FileResponse:
    return FileResponse(DASHBOARD_PATH)


@app.get("/healthz")
async def healthz() -> JSONResponse:
    return JSONResponse({"status": "ok"})


@app.websocket("/")
async def option_chain_socket(websocket: WebSocket) -> None:
    await websocket.accept()
    state = websocket.app.state.option_chain
    adapter = FastAPIWebSocketAdapter(websocket)
    await client_handler(adapter, state["initial_payload"], state["plan"], state["context"])
