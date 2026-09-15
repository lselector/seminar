#!/usr/bin/env python3
"""
Screenshot one block of a web page: the one around a heading.

Chrome's --screenshot flag can only capture the top of a
page. This module drives headless Chrome over the DevTools
protocol instead. It opens the page, waits for an element
whose own text is the heading, walks up to the first
ancestor tall enough to hold what sits under the heading (a
chart, say), scrolls it into view so lazy charts draw, and
captures just that rectangle at twice the pixel density.
page_text reads a page's visible text the same way.

Usage:
    from sources.element_shot import shoot_element
    path = shoot_element(
        "https://artificialanalysis.ai/evaluations/"
        "artificial-analysis-intelligence-index",
        "Cost per Intelligence Index Task", "/tmp/chart.png")

Created: 2026-09-14
Last updated: 2026-09-14
"""

import base64
import contextlib
import json
import os
import shutil
import socket
import subprocess
import tempfile
import time
import urllib.request

import websocket

from sources import fetch_images as F

WIDTH, HEIGHT = 1600, 1000
SCALE = 2
MIN_HEIGHT = 300
PAD = 12
LOAD_TIMEOUT = 60
SETTLE = 4
ERRORS = (OSError, RuntimeError, KeyError, TypeError,
          websocket.WebSocketException)

# Finds the heading, keeps its block in window.__shot and
# scrolls it into view. Returns false until the heading is
# on the page.
LOCATE_JS = """(() => {
  const want = %s.toLowerCase();
  const hits = [...document.querySelectorAll('body *')].filter(
    e => e.childElementCount === 0 &&
         e.textContent.trim().toLowerCase() === want);
  if (!hits.length) return false;
  let el = hits[0];
  while (el.parentElement &&
         el.getBoundingClientRect().height < %d) {
    el = el.parentElement;
  }
  window.__shot = el;
  el.scrollIntoView({block: 'center'});
  return true;
})()"""

MEASURE_JS = """(() => {
  const r = window.__shot.getBoundingClientRect();
  return {x: r.x + scrollX, y: r.y + scrollY,
          width: r.width, height: r.height};
})()"""


class Cdp:
    """A minimal DevTools protocol client for one tab."""

    # --------------------------------------
    def __init__(self, ws):
        """Wrap an open websocket to a page target."""
        self.ws = ws
        self.last = 0

    # --------------------------------------
    def call(self, method, **params):
        """Send one command and wait for its reply."""
        self.last += 1
        self.ws.send(json.dumps({"id": self.last,
                                 "method": method,
                                 "params": params}))
        while True:
            reply = json.loads(self.ws.recv())
            if reply.get("id") != self.last:
                continue
            if "error" in reply:
                raise RuntimeError(reply["error"].get("message"))
            return reply.get("result", {})

    # --------------------------------------
    def js(self, expression):
        """Evaluate JavaScript and return its value."""
        result = self.call("Runtime.evaluate",
                           expression=expression,
                           returnByValue=True)
        return result.get("result", {}).get("value")


# --------------------------------------------------------------
def free_port():
    """A local TCP port nobody is listening on."""
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


# --------------------------------------------------------------
def launch(port, profile, visible=False):
    """Start Chrome with remote debugging on.

    Headless by default. visible opens a real window, parked
    off-screen, for sites whose bot check stops headless
    Chrome (Cloudflare on trueup.io).
    """
    mode = (["--window-position=-3000,0", "--no-first-run",
             "--no-default-browser-check"] if visible
            else ["--headless=new", "--disable-gpu"])
    return subprocess.Popen(
        [F.CHROME, *mode, "--hide-scrollbars",
         f"--remote-debugging-port={port}",
         "--remote-allow-origins=*",
         f"--user-data-dir={profile}",
         f"--window-size={WIDTH},{HEIGHT}", "about:blank"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# --------------------------------------------------------------
@contextlib.contextmanager
def browser(visible=False):
    """A DevTools session on a fresh Chrome, always closed."""
    port, profile = free_port(), tempfile.mkdtemp()
    chrome = launch(port, profile, visible)
    try:
        yield connect(port)
    finally:
        chrome.kill()
        chrome.wait()
        shutil.rmtree(profile, ignore_errors=True)


# --------------------------------------------------------------
def connect(port):
    """Open a DevTools session on Chrome's first tab."""
    deadline = time.time() + 20
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(
                    f"http://127.0.0.1:{port}/json") as reply:
                tabs = json.load(reply)
            page = next(t for t in tabs if t["type"] == "page")
            return Cdp(websocket.create_connection(
                page["webSocketDebuggerUrl"], timeout=60))
        except (OSError, StopIteration):
            time.sleep(0.3)
    raise RuntimeError("Chrome did not start")


# --------------------------------------------------------------
def open_page(cdp, url, width=WIDTH):
    """Load a page at a given viewport width."""
    cdp.call("Emulation.setDeviceMetricsOverride", width=width,
             height=HEIGHT, deviceScaleFactor=SCALE,
             mobile=False)
    cdp.call("Page.enable")
    cdp.call("Page.navigate", url=url)


# --------------------------------------------------------------
def wait_for(cdp, expression, what):
    """Poll a JavaScript expression until it is truthy."""
    deadline = time.time() + LOAD_TIMEOUT
    while True:
        value = cdp.js(expression)
        if value:
            return value
        if time.time() > deadline:
            raise RuntimeError(f"{what} not found")
        time.sleep(1)


# --------------------------------------------------------------
def find_block(cdp, url, heading, width=WIDTH, ready=None):
    """Load the page and return the block's page rectangle.

    ready is an optional JavaScript expression that is true
    once the block has its data, for charts that load it
    after the page.
    """
    open_page(cdp, url, width)
    locate = LOCATE_JS % (json.dumps(heading), MIN_HEIGHT)
    wait_for(cdp, locate, f"heading {heading!r}")
    if ready:
        wait_for(cdp, ready, "chart data")
    time.sleep(SETTLE)
    cdp.js(locate)
    return cdp.js(MEASURE_JS)


# --------------------------------------------------------------
def capture(cdp, rect, path):
    """Save one page rectangle, with a little margin, as PNG."""
    clip = {"x": max(0, rect["x"] - PAD),
            "y": max(0, rect["y"] - PAD),
            "width": rect["width"] + 2 * PAD,
            "height": rect["height"] + 2 * PAD, "scale": 1}
    shot = cdp.call("Page.captureScreenshot", format="png",
                    clip=clip, captureBeyondViewport=True)
    with open(path, "wb") as handle:
        handle.write(base64.b64decode(shot["data"]))
    return path


# --------------------------------------------------------------
def shoot_element(url, heading, path, width=WIDTH,
                  visible=False, ready=None):
    """Screenshot the block under a heading; path or None."""
    if not os.path.isfile(F.CHROME):
        F.log_message(f"  Chrome not found at {F.CHROME}")
        return None
    try:
        with browser(visible) as cdp:
            rect = find_block(cdp, url, heading, width, ready)
            return capture(cdp, rect, path)
    except ERRORS as exc:
        F.log_message(f"  could not shoot {heading!r}: {exc}")
        return None


# --------------------------------------------------------------
def page_text(url, marker, visible=False):
    """The page's visible text once marker shows; or None."""
    if not os.path.isfile(F.CHROME):
        F.log_message(f"  Chrome not found at {F.CHROME}")
        return None
    probe = ("document.body && document.body.innerText"
             f".includes({json.dumps(marker)})")
    try:
        with browser(visible) as cdp:
            open_page(cdp, url)
            wait_for(cdp, probe, repr(marker))
            return cdp.js("document.body.innerText")
    except ERRORS as exc:
        F.log_message(f"  could not read {url}: {exc}")
        return None


# --------------------------------------------------------------
if __name__ == "__main__":
    import sys
    print(shoot_element(sys.argv[1], sys.argv[2], sys.argv[3]))
