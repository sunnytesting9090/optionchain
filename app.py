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

# =========================================================
# PATHS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_PATH = BASE_DIR / "option_chain.html"


# =========================================================
# WEBSOCKET ADAPTER
# =========================================================

class FastAPIWebSocketAdapter:
    def __init__(self, websocket: WebSocket) -> None:
        self.websocket = websocket

    async def send(self, message: str) -> None:
        await self.websocket.send_text(message)

    def __aiter__(self):
        return self

    async def __anext__(self):
        try:
            return await self.websocket.receive_text()
        except WebSocketDisconnect:
            raise StopAsyncIteration


# =========================================================
# BUILD APPLICATION STATE
# =========================================================

def build_app_state() -> dict[str, Any]:

    log("Loading Zerodha access token...")
    access_token = load_access_token()

    log("Building Kite client...")
    kite = build_kite_client(access_token)

    log("Loading instruments...")
    instrument_frame = load_instrument_universe(kite)

    log("Building runtime plan...")
    plan = build_runtime_plan(kite, instrument_frame)

    log("Seeding initial ticks...")
    latest_ticks = seed_initial_ticks(kite, plan)

    metadata = token_metadata(plan)
    expiry_map = build_expiry_map(instrument_frame)
    initial_payload = initial_messages(latest_ticks, metadata)

    tick_queue: queue.Queue[list[dict[str, Any]]] = queue.Queue()

    stop_event = threading.Event()

    ticker_ws_holder: dict[str, Any] = {}

    log("Starting Kite WebSocket stream...")

    ticker = start_kite_stream(
        access_token,
        plan,
        tick_queue,
        stop_event,
        ticker_ws_holder,
    )

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


# =========================================================
# FASTAPI LIFESPAN
# =========================================================

@asynccontextmanager
async def lifespan(app: FastAPI):

    log("Initializing option-chain app state...")

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

        log("Shutting down application...")

        state["stop_event"].set()

        broadcaster.cancel()

        try:
            await broadcaster
        except asyncio.CancelledError:
            pass

        try:
            state["ticker"].close()
        except Exception as e:
            log(f"Ticker close error: {e}")

        CLIENTS.clear()

        log("Shutdown complete.")


# =========================================================
# FASTAPI APP
# =========================================================

app = FastAPI(
    title="Zerodha Option Chain",
    lifespan=lifespan,
)


# =========================================================
# ROUTES
# =========================================================

@app.get("/")
async def dashboard():

    if DASHBOARD_PATH.exists():
        return FileResponse(DASHBOARD_PATH)

    return JSONResponse(
        {
            "status": "running",
            "message": "Dashboard HTML not found."
        }
    )


@app.get("/healthz")
async def healthz():

    return JSONResponse(
        {
            "status": "ok"
        }
    )


# =========================================================
# WEBSOCKET ROUTE
# =========================================================

@app.websocket("/ws")
async def option_chain_socket(websocket: WebSocket):

    await websocket.accept()

    state = websocket.app.state.option_chain

    adapter = FastAPIWebSocketAdapter(websocket)

    try:

        await client_handler(
            adapter,
            state["initial_payload"],
            state["plan"],
            state["context"],
        )

    except WebSocketDisconnect:

        log("WebSocket client disconnected.")

    except Exception as e:

        log(f"WebSocket error: {e}")


# =========================================================
# LOCAL DEVELOPMENT ENTRY
# =========================================================

if __name__ == "__main__":

    import uvicorn

    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
    )