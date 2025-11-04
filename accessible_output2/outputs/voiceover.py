from __future__ import absolute_import

import ctypes
import logging
import os
from typing import Optional, Union

try:
    import wx  # type: ignore
except ImportError:  # pragma: no cover - wx will always be available in Flux
    wx = None

from .base import Output

logger = logging.getLogger(__name__)


class VoiceOverBridge:
    """
    Thin ctypes wrapper around the libVoiceOver bridge compiled from VoiceOver.m.
    """

    def __init__(
        self,
        dylib_name: str = "libVoiceOver.dylib",
        lib: Optional[ctypes.CDLL] = None,
    ):
        self.initialized = False
        if lib is None:
            self._path = self._resolve_dylib_path(dylib_name)
            self.lib = ctypes.CDLL(self._path)
        else:
            self._path = getattr(lib, "__file__", "<in-memory>")
            self.lib = lib
        self._configure_signatures()

    @staticmethod
    def _resolve_dylib_path(dylib_name: str) -> str:
        candidates = []

        # Package-local lib/ directory when running from source.
        package_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        candidates.append(os.path.join(package_root, "lib", dylib_name))

        # Application bundle or frozen binaries.
        try:
            from platform_utils import paths  # type: ignore

            if paths.is_frozen():
                candidates.append(
                    os.path.join(
                        paths.embedded_data_path(),
                        "accessible_output2",
                        "lib",
                        dylib_name,
                    )
                )
            else:
                candidates.append(
                    os.path.join(paths.module_path(), "lib", dylib_name)
                )
            candidates.append(
                os.path.join(
                    paths.embedded_data_path(),
                    "lib",
                    "accessible_output2",
                    "lib",
                    dylib_name,
                )
            )
        except Exception:
            # platform_utils is optional when running outside frozen builds.
            pass

        # Absolute path relative to current working directory as a final fallback.
        candidates.append(os.path.abspath(dylib_name))

        searched = []
        for candidate in candidates:
            if not candidate:
                continue
            candidate = os.path.abspath(candidate)
            if candidate in searched:
                continue
            searched.append(candidate)
            if os.path.exists(candidate):
                return candidate

        raise FileNotFoundError(
            "VoiceOver dylib not found. Searched: %s" % ", ".join(searched)
        )

    def _configure_signatures(self) -> None:
        required = ("vo_init_with_window", "vo_is_running", "vo_announce", "vo_shutdown")
        for name in required:
            if not hasattr(self.lib, name):
                raise AttributeError("Loaded VoiceOver library missing symbol: %s" % name)

        try:
            self.lib.vo_init_with_window.argtypes = [ctypes.c_void_p]
            self.lib.vo_init_with_window.restype = ctypes.c_bool
            self.lib.vo_is_running.argtypes = []
            self.lib.vo_is_running.restype = ctypes.c_bool
            self.lib.vo_announce.argtypes = [ctypes.c_char_p, ctypes.c_int]
            self.lib.vo_announce.restype = ctypes.c_bool
            self.lib.vo_shutdown.argtypes = []
            self.lib.vo_shutdown.restype = None
        except AttributeError:
            # Allow substitutes used during testing that do not expose ctypes metadata.
            pass

    def init(self, hwnd: Union[int, ctypes.c_void_p]) -> bool:
        if hwnd in (None, 0):
            raise ValueError("VoiceOver init requires a valid window handle")

        ptr = hwnd if isinstance(hwnd, ctypes.c_void_p) else ctypes.c_void_p(int(hwnd))
        ok = self.lib.vo_init_with_window(ptr)
        self.initialized = bool(ok)
        return self.initialized

    def is_running(self) -> bool:
        return bool(self.lib.vo_is_running())

    def speak(self, text: str, interrupt: bool = True) -> bool:
        if not self.initialized:
            raise RuntimeError("VoiceOver not initialized with a window yet.")
        if not isinstance(text, str):
            raise TypeError("VoiceOver.speak expects a string message")

        payload = text.encode("utf-8")
        return bool(self.lib.vo_announce(payload, 1 if interrupt else 0))

    def shutdown(self) -> None:
        self.lib.vo_shutdown()
        self.initialized = False


class VoiceOver(Output):
    """Speech output that bridges to Apple's VoiceOver via libVoiceOver.dylib."""

    name = "VoiceOver"
    priority = 90
    system_output = True

    def __init__(self, *args, **kwargs):
        super().__init__()
        self._bridge: Optional[VoiceOverBridge] = None
        self._window_handle: Optional[ctypes.c_void_p] = None
        self._initialization_attempted = False
        self._initialized = False
        self._last_error: Optional[Exception] = None

    def set_main_window(self, window) -> None:
        """Allow the application to provide an explicit wx window handle."""
        handle = self._extract_handle(window)
        if handle:
            self._window_handle = handle
            self._initialization_attempted = False
            self._initialized = False
            if self._bridge is not None:
                try:
                    self._bridge.shutdown()
                except Exception:
                    pass
                self._bridge = None

    def _ensure_bridge(self) -> bool:
        if self._initialized and self._bridge:
            return True

        if self._bridge is None:
            try:
                self._bridge = VoiceOverBridge()
            except Exception as exc:
                self._last_error = exc
                logger.error("Failed to construct VoiceOver bridge: %s", exc)
                return False

        handle = self._window_handle or self._auto_detect_window_handle()
        if not handle:
            if not self._initialization_attempted:
                logger.debug("VoiceOver window handle unavailable; deferring initialization")
            self._initialization_attempted = True
            return False

        try:
            if self._bridge.init(handle):
                self._initialized = True
                self._initialization_attempted = True
                return True
            logger.error("VoiceOver bridge initialization returned False")
        except Exception as exc:
            self._last_error = exc
            logger.error("VoiceOver bridge initialization failed: %s", exc)
        self._initialization_attempted = True
        return False

    def _auto_detect_window_handle(self) -> Optional[ctypes.c_void_p]:
        if wx is None:
            return None
        app = wx.GetApp()
        if app is None:
            return None
        window = app.GetTopWindow()
        if window is None:
            return None
        handle = self._extract_handle(window)
        if handle:
            self._window_handle = handle
        return handle

    @staticmethod
    def _extract_handle(window) -> Optional[ctypes.c_void_p]:
        if window is None:
            return None

        handle = None
        if hasattr(window, "MacGetTopLevelWindowRef"):
            try:
                handle = window.MacGetTopLevelWindowRef()
            except Exception:
                handle = None
        elif hasattr(window, "GetHandle"):
            try:
                handle = window.GetHandle()
            except Exception:
                handle = None

        if handle in (None, 0):
            return None

        if isinstance(handle, ctypes.c_void_p):
            return handle

        try:
            return ctypes.c_void_p(int(handle))
        except (TypeError, ValueError):
            logger.error("Could not convert wx window handle %r to pointer", handle)
            return None

    def speak(self, text, interrupt=False):
        if not self._ensure_bridge():
            return False
        try:
            return self._bridge.speak(text, bool(interrupt))
        except Exception as exc:
            logger.error("VoiceOver speak failed: %s", exc)
            return False

    def silence(self):
        if not self._bridge or not getattr(self._bridge, "initialized", False):
            return
        try:
            self._bridge.speak("", True)
        except Exception as exc:
            logger.debug("VoiceOver silence failed: %s", exc)

    def is_active(self):
        if not self._bridge:
            # Attempt initialization to verify VoiceOver state
            self._ensure_bridge()
        if not self._bridge:
            return False
        try:
            return self._bridge.is_running()
        except Exception as exc:
            logger.debug("VoiceOver is_running check failed: %s", exc)
            return False

    def shutdown(self):
        if self._bridge:
            try:
                self._bridge.shutdown()
            except Exception:
                pass
            finally:
                self._bridge = None
                self._initialized = False


output_class = VoiceOver
