from __future__ import annotations

import argparse
import asyncio
import json
import queue
import signal
import threading
import time
from datetime import datetime, time as day_time
from typing import Any

import websockets
from kiteconnect import KiteTicker

import config
from update_simple_option_chain import (
    INDEX_ORDER,
    SUPPORTED_INDICES,
    build_kite_client,
    build_expiry_map,
    build_runtime_plan,
    format_timestamp,
    get_best_depth_entry,
    load_access_token,
    load_instrument_universe,
    log,
    number_or_none,
    seed_initial_ticks,
)


CLIENTS: set[Any] = set()
BROADCAST_SECONDS = 0.12
ADVANCE_DECLINE_SECONDS = 5
OPTION_BATCH_SIZE = 450

CONSTITUENTS = {
    "NIFTY": [
        "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK", "BAJAJ-AUTO", "BAJFINANCE",
        "BAJAJFINSV", "BEL", "BHARTIARTL", "CIPLA", "COALINDIA", "DRREDDY", "EICHERMOT", "ETERNAL",
        "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE", "HEROMOTOCO", "HINDALCO", "HINDUNILVR", "ICICIBANK",
        "INDUSINDBK", "INFY", "ITC", "JIOFIN", "JSWSTEEL", "KOTAKBANK", "LT", "M&M", "MARUTI",
        "NESTLEIND", "NTPC", "ONGC", "POWERGRID", "RELIANCE", "SBILIFE", "SHRIRAMFIN", "SBIN",
        "SUNPHARMA", "TCS", "TATACONSUM", "TATAMOTORS", "TATASTEEL", "TECHM", "TITAN", "TRENT",
        "ULTRACEMCO", "WIPRO",
    ],
    "BANKNIFTY": [
        "AUBANK", "AXISBANK", "BANDHANBNK", "BANKBARODA", "CANBK", "FEDERALBNK", "HDFCBANK",
        "ICICIBANK", "IDFCFIRSTB", "INDUSINDBK", "KOTAKBANK", "PNB", "SBIN",
    ],
    "SENSEX": [
        "ASIANPAINT", "AXISBANK", "BAJFINANCE", "BAJAJFINSV", "BHARTIARTL", "HCLTECH", "HDFCBANK",
        "HINDUNILVR", "ICICIBANK", "INDUSINDBK", "INFY", "ITC", "KOTAKBANK", "LT", "M&M", "MARUTI",
        "NESTLEIND", "NTPC", "POWERGRID", "RELIANCE", "SBIN", "SUNPHARMA", "TCS", "TATASTEEL",
        "TECHM", "TITAN", "ULTRACEMCO", "ZOMATO",
    ],
}


def token_metadata(plan: dict[str, Any]) -> dict[int, dict[str, Any]]:
    metadata: dict[int, dict[str, Any]] = {}
    for index_name in INDEX_ORDER:
        index_plan = plan["indices"][index_name]
        metadata[int(index_plan["spot_token"])] = {
            "symbol": index_name,
            "kind": "spot",
            "spot_symbol": index_plan["spot_symbol"],
            "expiry": index_plan["expiry"],
        }

        paired = index_plan["paired"]
        for row in paired.itertuples(index=False):
            strike = float(row.strike)
            metadata[int(row.ce_token)] = {
                "symbol": index_name,
                "kind": "option",
                "expiry": index_plan["expiry"],
                "strike": strike,
                "type": "CE",
            }
            metadata[int(row.pe_token)] = {
                "symbol": index_name,
                "kind": "option",
                "expiry": index_plan["expiry"],
                "strike": strike,
                "type": "PE",
            }
    return metadata


def quote_key(exchange: str, tradingsymbol: str) -> str:
    return f"{exchange}:{tradingsymbol}"


def fetch_ltp_in_batches(kite: Any, keys: list[str]) -> dict[Any, Any]:
    quotes: dict[Any, Any] = {}
    for start_index in range(0, len(keys), OPTION_BATCH_SIZE):
        batch = keys[start_index : start_index + OPTION_BATCH_SIZE]
        if batch:
            quotes.update(kite.ltp(batch))
    return quotes


def build_chain_subscription(
    kite: Any,
    instrument_frame: Any,
    index_name: str,
    expiry: str,
) -> tuple[list[int], dict[int, dict[str, Any]], list[dict[str, Any]]]:
    meta = SUPPORTED_INDICES[index_name]
    filtered = instrument_frame[
        (instrument_frame["name"] == meta["name"])
        & (instrument_frame["segment"] == meta["segment"])
        & (instrument_frame["expiry"] == expiry)
        & (instrument_frame["instrument_type"].isin(["CE", "PE"]))
    ].copy()
    if filtered.empty:
        raise RuntimeError(f"No option contracts found for {index_name} {expiry}.")

    filtered["instrument_token"] = filtered["instrument_token"].astype(int)
    tokens = filtered["instrument_token"].tolist()
    metadata: dict[int, dict[str, Any]] = {}
    quote_keys: list[str] = []
    key_by_token: dict[int, str] = {}
    for row in filtered.itertuples(index=False):
        token = int(row.instrument_token)
        key = quote_key(row.exchange, row.tradingsymbol)
        quote_keys.append(key)
        key_by_token[token] = key
        metadata[token] = {
            "symbol": index_name,
            "kind": "option",
            "expiry": expiry,
            "strike": float(row.strike),
            "type": row.instrument_type,
        }

    spot_quote = kite.quote([meta["spot_symbol"]]).get(meta["spot_symbol"], {})
    spot_token = int(spot_quote.get("instrument_token"))
    tokens.append(spot_token)
    metadata[spot_token] = {
        "symbol": index_name,
        "kind": "spot",
        "spot_symbol": meta["spot_symbol"],
        "expiry": expiry,
    }

    initial: list[dict[str, Any]] = [
        {
            "symbol": index_name,
            "spot": number_or_none(spot_quote.get("last_price")),
            "time": format_timestamp(spot_quote.get("timestamp") or spot_quote.get("exchange_timestamp")),
        }
    ]
    quotes = fetch_ltp_in_batches(kite, quote_keys)
    for token, key in key_by_token.items():
        quote = quotes.get(key) or {}
        contract_meta = metadata[token]
        initial.append(
            {
                "symbol": index_name,
                "expiry": expiry,
                "strike": contract_meta["strike"],
                "type": contract_meta["type"],
                "ltp": number_or_none(quote.get("last_price")),
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

    return tokens, metadata, initial


def find_option_token(instrument_frame: Any, index_name: str, expiry: str, strike: float, option_type: str) -> int:
    meta = SUPPORTED_INDICES[index_name]
    filtered = instrument_frame[
        (instrument_frame["name"] == meta["name"])
        & (instrument_frame["segment"] == meta["segment"])
        & (instrument_frame["expiry"] == expiry)
        & (instrument_frame["instrument_type"] == option_type)
        & (instrument_frame["strike"].astype(float) == float(strike))
    ]
    if filtered.empty:
        raise RuntimeError(f"No {option_type} token found for {index_name} {expiry} {strike}.")
    return int(filtered.iloc[0]["instrument_token"])


def build_strike_history_payload(
    kite: Any,
    instrument_frame: Any,
    index_name: str,
    expiry: str,
    call_strike: float,
    put_strike: float,
) -> dict[str, Any]:
    start = datetime.combine(datetime.now().date(), day_time(9, 15)).strftime("%Y-%m-%d %H:%M:%S")
    end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    legs = {
        "CE": find_option_token(instrument_frame, index_name, expiry, call_strike, "CE"),
        "PE": find_option_token(instrument_frame, index_name, expiry, put_strike, "PE"),
    }
    history: dict[str, list[dict[str, Any]]] = {}
    for option_type, token in legs.items():
        candles = kite.historical_data(token, start, end, "minute", oi=True)
        history[option_type] = [
            {
                "time": format_timestamp(row.get("date")),
                "price": number_or_none(row.get("close")),
                "oi": row.get("oi"),
            }
            for row in candles
        ]
    return {
        "type": "strike_history",
        "symbol": index_name,
        "expiry": expiry,
        "call_strike": call_strike,
        "put_strike": put_strike,
        "rows": {option_type: len(rows) for option_type, rows in history.items()},
        "history": history,
    }


def tick_timestamp(tick: dict[str, Any]) -> str:
    value = tick.get("exchange_timestamp") or tick.get("last_trade_time") or datetime.now()
    return format_timestamp(value) or datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_tick(tick: dict[str, Any], metadata: dict[int, dict[str, Any]]) -> dict[str, Any] | None:
    token = int(tick.get("instrument_token", 0) or 0)
    meta = metadata.get(token)
    if not meta:
        return None

    if meta["kind"] == "spot":
        return {
            "symbol": meta["symbol"],
            "spot": number_or_none(tick.get("last_price")),
            "time": tick_timestamp(tick),
        }

    ohlc = tick.get("ohlc") or {}
    best_buy = get_best_depth_entry(tick, "buy")
    return {
        "symbol": meta["symbol"],
        "expiry": meta["expiry"],
        "strike": meta["strike"],
        "type": meta["type"],
        "ltp": number_or_none(tick.get("last_price")),
        "open": number_or_none(ohlc.get("open")),
        "previous_close": number_or_none(ohlc.get("close")),
        "bid": number_or_none(best_buy.get("price")),
        "volume": tick.get("volume_traded"),
        "oi": tick.get("oi"),
        "time": tick_timestamp(tick),
    }


def initial_messages(latest_ticks: dict[int, dict[str, Any]], metadata: dict[int, dict[str, Any]]) -> list[dict[str, Any]]:
    messages: list[dict[str, Any]] = []
    for token, tick in latest_ticks.items():
        normalized = normalize_tick(tick, metadata)
        if normalized is not None:
            messages.append(normalized)
    return messages


def build_advance_decline_payloads(kite: Any) -> list[dict[str, Any]]:
    payloads: list[dict[str, Any]] = []
    all_symbols = sorted({symbol for symbols in CONSTITUENTS.values() for symbol in symbols})
    quotes: dict[str, Any] = {}
    for start_index in range(0, len(all_symbols), 40):
        keys = [f"NSE:{symbol}" for symbol in all_symbols[start_index : start_index + 40]]
        try:
            quotes.update(kite.quote(keys))
        except Exception:
            for key in keys:
                try:
                    quotes.update(kite.quote([key]))
                except Exception:
                    continue
    by_symbol: dict[str, dict[str, Any]] = {}
    for key, quote in quotes.items():
        symbol = key.split(":", 1)[-1]
        close = number_or_none((quote.get("ohlc") or {}).get("close"))
        last_price = number_or_none(quote.get("last_price"))
        change = None
        if last_price is not None and close not in (None, 0):
            change = last_price - close
        by_symbol[symbol] = {
            "symbol": symbol,
            "last_price": last_price,
            "change": change,
        }

    for index_name, symbols in CONSTITUENTS.items():
        items = [by_symbol[symbol] for symbol in symbols if symbol in by_symbol]
        payloads.append(
            {
                "type": "advance_decline",
                "symbol": index_name,
                "items": items,
                "message": (
                    f"Advance/decline loaded {len(items)} of {len(symbols)} symbols for {index_name}."
                    if items
                    else f"No advance/decline quotes returned for {index_name}."
                ),
            }
        )
    return payloads


def start_kite_stream(
    access_token: str,
    plan: dict[str, Any],
    outbound_queue: queue.Queue[list[dict[str, Any]]],
    stop_event: threading.Event,
    ticker_ws_holder: dict[str, Any],
) -> KiteTicker:
    ticker = KiteTicker(
        config.cred["api_key"],
        access_token,
        reconnect=True,
        reconnect_max_tries=50,
        reconnect_max_delay=60,
    )

    connected_event = threading.Event()

    def on_connect(ws, response):
        ticker_ws_holder["ws"] = ws
        log("Connected to Zerodha ticker. Subscribing to option chain tokens.")
        ws.subscribe(plan["subscription_tokens"])
        if plan["spot_tokens"]:
            ws.set_mode(ws.MODE_QUOTE, plan["spot_tokens"])
        if plan["option_tokens"]:
            ws.set_mode(ws.MODE_FULL, plan["option_tokens"])
        connected_event.set()

    def on_ticks(ws, ticks):
        if ticks:
            outbound_queue.put(ticks)

    def on_close(ws, code, reason):
        log(f"Zerodha ticker closed: {code} {reason}")
        stop_event.set()

    def on_error(ws, code, reason):
        log(f"Zerodha ticker error {code}: {reason}")

    ticker.on_connect = on_connect
    ticker.on_ticks = on_ticks
    ticker.on_close = on_close
    ticker.on_error = on_error

    ticker.connect(threaded=True)
    if not connected_event.wait(timeout=20):
        raise RuntimeError("Zerodha WebSocket did not connect within 20 seconds.")
    return ticker


async def send_json(websocket: Any, payload: Any) -> None:
    await websocket.send(json.dumps(payload, separators=(",", ":"), default=str))


async def client_handler(websocket: Any, initial_payload: list[dict[str, Any]], plan: dict[str, Any], context: dict[str, Any]) -> None:
    CLIENTS.add(websocket)
    try:
        await send_json(
            websocket,
            {
                "type": "hello",
                "indices": {
                    index_name: {
                        "expiry": plan["indices"][index_name]["expiry"],
                        "expiries": context["expiry_map"].get(index_name, []),
                        "spot_symbol": plan["indices"][index_name]["spot_symbol"],
                    }
                    for index_name in INDEX_ORDER
                },
            },
        )
        try:
            for payload in build_advance_decline_payloads(context["kite"]):
                await send_json(websocket, payload)
        except Exception as exc:
            log(f"Initial advance/decline quote refresh failed: {exc}")
        for start in range(0, len(initial_payload), 400):
            await send_json(websocket, initial_payload[start : start + 400])
        async for message in websocket:
            try:
                request = json.loads(message)
            except json.JSONDecodeError:
                continue
            if request.get("action") == "subscribe":
                index_name = request.get("symbol")
                expiry = request.get("expiry")
                if index_name not in INDEX_ORDER or not expiry:
                    continue
                try:
                    tokens, new_metadata, initial = build_chain_subscription(
                        context["kite"],
                        context["instrument_frame"],
                        index_name,
                        expiry,
                    )
                    context["metadata"].update(new_metadata)
                    context["latest_expiry"][(index_name, expiry)] = True
                    ticker_ws = context["ticker_ws_holder"].get("ws")
                    if ticker_ws is not None:
                        ticker_ws.subscribe(tokens)
                        spot_tokens = [token for token in tokens if new_metadata[token]["kind"] == "spot"]
                        option_tokens = [token for token in tokens if new_metadata[token]["kind"] == "option"]
                        if spot_tokens:
                            ticker_ws.set_mode(ticker_ws.MODE_QUOTE, spot_tokens)
                        if option_tokens:
                            ticker_ws.set_mode(ticker_ws.MODE_FULL, option_tokens)
                    await send_json(websocket, initial)
                    log(f"Browser subscribed to {index_name} {expiry}.")
                except Exception as exc:
                    await send_json(websocket, {"type": "error", "message": str(exc)})
            elif request.get("action") == "history":
                try:
                    payload = build_strike_history_payload(
                        context["kite"],
                        context["instrument_frame"],
                        request["symbol"],
                        request["expiry"],
                        float(request["call_strike"]),
                        float(request["put_strike"]),
                    )
                    await send_json(websocket, payload)
                except Exception as exc:
                    await send_json(websocket, {"type": "error", "message": str(exc)})
    finally:
        CLIENTS.discard(websocket)


async def broadcast_loop(
    kite: Any,
    tick_queue: queue.Queue[list[dict[str, Any]]],
    latest_ticks: dict[int, dict[str, Any]],
    metadata: dict[int, dict[str, Any]],
    stop_event: threading.Event,
) -> None:
    next_advance_decline = 0.0
    while not stop_event.is_set():
        started = time.monotonic()
        normalized_batch: list[dict[str, Any]] = []

        while True:
            try:
                ticks = tick_queue.get_nowait()
            except queue.Empty:
                break
            for tick in ticks:
                token = int(tick.get("instrument_token", 0) or 0)
                if token:
                    latest_ticks[token] = tick
                normalized = normalize_tick(tick, metadata)
                if normalized is not None:
                    normalized_batch.append(normalized)

        if time.monotonic() >= next_advance_decline:
            try:
                normalized_batch.extend(build_advance_decline_payloads(kite))
            except Exception as exc:
                log(f"Advance/decline quote refresh failed: {exc}")
            next_advance_decline = time.monotonic() + ADVANCE_DECLINE_SECONDS

        if normalized_batch and CLIENTS:
            payload = json.dumps(normalized_batch, separators=(",", ":"), default=str)
            dead_clients = []
            for websocket in list(CLIENTS):
                try:
                    await websocket.send(payload)
                except Exception:
                    dead_clients.append(websocket)
            for websocket in dead_clients:
                CLIENTS.discard(websocket)

        elapsed = time.monotonic() - started
        await asyncio.sleep(max(0.01, BROADCAST_SECONDS - elapsed))


async def run_server(host: str, port: int) -> None:
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

    loop = asyncio.get_running_loop()
    for signame in ("SIGINT", "SIGTERM"):
        if hasattr(signal, signame):
            try:
                loop.add_signal_handler(getattr(signal, signame), stop_event.set)
            except NotImplementedError:
                pass

    log(f"Browser WebSocket server ready at ws://{host}:{port}")
    async with websockets.serve(lambda ws: client_handler(ws, initial_payload, plan, context), host, port):
        broadcaster = asyncio.create_task(broadcast_loop(kite, tick_queue, latest_ticks, metadata, stop_event))
        try:
            while not stop_event.is_set():
                await asyncio.sleep(0.25)
        finally:
            broadcaster.cancel()
            ticker.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Local Zerodha option-chain WebSocket bridge for option_chain.html.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    asyncio.run(run_server(args.host, args.port))


if __name__ == "__main__":
    main()
