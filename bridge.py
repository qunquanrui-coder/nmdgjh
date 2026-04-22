# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import threading
from typing import Any, Callable, Dict, Optional

_exposed_functions: Dict[str, Callable[..., Any]] = {}

_window = None
_window_lock = threading.Lock()


def expose(func: Optional[Callable[..., Any]] = None, name: Optional[str] = None):
    def decorator(f: Callable[..., Any]):
        exposed_name = name or f.__name__
        _exposed_functions[exposed_name] = f
        return f

    if func is not None and callable(func):
        return decorator(func)

    return decorator


def set_window(window) -> None:
    global _window
    with _window_lock:
        _window = window


def get_window():
    with _window_lock:
        return _window


def call_frontend(function_name: str, *args: Any):
    window = get_window()
    if window is None:
        return None

    fn_name = json.dumps(function_name, ensure_ascii=False)
    payload = json.dumps(list(args), ensure_ascii=False)

    script = f"""
    (function() {{
        const fn = window[{fn_name}];
        if (typeof fn === 'function') {{
            return fn(...{payload});
        }}
        return null;
    }})()
    """

    try:
        return window.evaluate_js(script)
    except Exception:
        return None


def update_terminal(message: str):
    return call_frontend("update_terminal", message)
