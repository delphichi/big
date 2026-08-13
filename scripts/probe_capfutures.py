"""一次性探針: 抓 CTBC 期貨 (capfutures) ETF SPA 頁面 + JS bundle, 找 API URL."""
import re, sys, requests

BASE = "https://etf.capfutures.com"
URL = f"{BASE}/?view=holdings&fund=00406A"
HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,*/*",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

def probe_endpoints(text, label):
    """從 HTML 或 JS bundle 找 API endpoint / 下載 URL."""
    patterns = [
        r'["\'`](https?://[^"\'`\s]+)["\'`]',
        r'["\'`](/[a-zA-Z][a-zA-Z0-9_\-./]{3,120})["\'`]',
        r'fetch\s*\(\s*["\'`]([^"\'`]+)["\'`]',
        r'axios\.(?:get|post)\s*\(\s*["\'`]([^"\'`]+)["\'`]',
    ]
    found = set()
    for p in patterns:
        for m in re.findall(p, text):
            u = m if isinstance(m, str) else m[0]
            lu = u.lower()
            if any(k in lu for k in ("fund", "etf", "holding", "asset", "pcf", "excel", "xls", "csv", "download", "api", "component", "constituent", "weight", "portfolio", "position", "capfutures", "406a", "00406", "券商", "持股", "成分")):
                found.add(u)
    print(f"[{label}: {len(found)} URLs matched]")
    for u in sorted(found):
        print(f"  {u}")
    return found

def main():
    # 1) 抓 SPA 主頁
    r = requests.get(URL, headers=HDR, timeout=30)
    print(f"[HTML STATUS] {r.status_code} len={len(r.content)}")
    if r.status_code != 200:
        print(r.text[:2000]); return 1
    scripts = re.findall(r"<script[^>]+src=['\"]([^'\"]+)['\"]", r.text, re.I)
    probe_endpoints(r.text, "HTML")

    # 2) 抓 JS bundle
    for s in scripts:
        js_url = s if s.startswith("http") else BASE + s
        print(f"\n=== JS: {js_url} ===")
        try:
            jr = requests.get(js_url, headers=HDR, timeout=60)
            print(f"[JS STATUS] {jr.status_code} len={len(jr.content)}")
            if jr.status_code != 200: continue
            js = jr.text
            probe_endpoints(js, "JS")
            # 額外: 找 "GET"/"POST" 附近的 URL 樣式
            for m in re.finditer(r'(?:GET|POST|PUT|DELETE)["\'`,\s]{1,10}["\'`](https?://[^"\'`]+|/[^"\'`]+)["\'`]', js):
                print(f"  [HTTP CALL] {m.group(0)[:200]}")
            # 找關鍵字上下文
            for kw in ("holding", "asset", "component", "portfolio", "constituent", "持股", "成分"):
                for m in re.finditer(kw, js, re.I):
                    ctx = js[max(0, m.start()-120):m.end()+120]
                    if any(u in ctx for u in ("http", "/api/", "url:", "path:")):
                        print(f"  [CTX '{kw}'] ...{ctx}...")
                        break  # 一個關鍵字只印一次
        except Exception as e:
            print(f"[JS FAIL] {e}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
