"""Portal-context API calls (axios first, fetch fallback) and token extraction."""

from __future__ import annotations

import json

from .constants import log


class SessionExpired(Exception):
    """Raised on HTTP 440 - portal session timed out."""


def get_auth_tokens(context) -> tuple[str | None, str | None]:
    """Extract XSRF token and tenant ID from browser cookies."""
    cookies = context.cookies()
    xsrf = next((c["value"] for c in cookies if c["name"] == "XSRF-TOKEN"), None)
    tid = next((c["value"] for c in cookies if c["name"] == "x-tid"), None)
    return xsrf, tid


def api_call(page, method: str, url: str, xsrf: str, tid: str | None) -> dict:
    """Make an authenticated API call from within the portal page.

    Tries axios first (the portal's request interceptors run on its instance),
    falls back to fetch with explicit headers. Returns the parsed JSON body.
    """
    result = page.evaluate(
        """async ({method, url, xsrf, tid}) => {
            if (typeof axios !== 'undefined') {
                try {
                    const opts = { timeout: 30000 };
                    const resp = method === 'POST'
                        ? await axios.post(url, {}, opts)
                        : await axios.get(url, opts);
                    return {
                        ok: resp.status >= 200 && resp.status < 300,
                        status: resp.status,
                        body: typeof resp.data === 'string' ? resp.data : JSON.stringify(resp.data),
                        via: 'axios'
                    };
                } catch (axErr) {
                    if (axErr.response) {
                        return {
                            ok: false, status: axErr.response.status,
                            body: typeof axErr.response.data === 'string'
                                ? axErr.response.data : JSON.stringify(axErr.response.data),
                            via: 'axios-error'
                        };
                    }
                }
            }
            const headers = { 'Accept': 'application/json', 'X-XSRF-TOKEN': xsrf };
            if (tid) headers['x-tid'] = tid;
            const opts = { method, headers, credentials: 'include' };
            try {
                const resp = await fetch(url, opts);
                return { ok: resp.ok, status: resp.status, body: await resp.text(), via: 'fetch' };
            } catch (e) {
                return { ok: false, status: 0, body: e.toString(), via: 'fetch-error' };
            }
        }""",
        {"method": method, "url": url, "xsrf": xsrf or "", "tid": tid or ""},
    )

    log.info("API %s %s -> %s (via %s)", method, url.split("/")[-1][:60],
             result.get("status"), result.get("via"))

    if not result["ok"]:
        status = result["status"]
        if status == 440:
            raise SessionExpired("440 Login Timeout")
        body_preview = (result.get("body") or "")[:500]
        raise RuntimeError(f"HTTP {status}: {body_preview}")

    return json.loads(result["body"])


def extract_list(data) -> list:
    """Pull a list of items from an API response (list or dict with a list value)."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("value", "items", "Aggregates"):
            if isinstance(data.get(key), list):
                return data[key]
        for v in data.values():
            if isinstance(v, list) and len(v) > 0:
                return v
    return []
