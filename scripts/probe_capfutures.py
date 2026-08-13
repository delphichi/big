"""一次性探針: 抓 CTBC 期貨 (capfutures) ETF 頁面, 找 API/下載 URL."""
import re, sys, requests

URL = "https://etf.capfutures.com/?view=holdings&fund=00406A"
HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,*/*",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

def main():
    try:
        r = requests.get(URL, headers=HDR, timeout=30)
    except Exception as e:
        print(f"[FAIL] GET: {e}"); return 1
    print(f"[STATUS] {r.status_code} len={len(r.content)}")
    if r.status_code != 200:
        print(r.text[:2000]); return 1
    html = r.text
    # 找 API / xhr / fetch / ajax / download / xlsx / csv / json endpoint
    patterns = [
        r"(?i)(?:url|href|src|action)\s*[:=]\s*['\"]([^'\"]{4,200})['\"]",
        r"(?i)fetch\s*\(\s*['\"]([^'\"]+)['\"]",
        r"(?i)ajax\s*\(\s*\{[^}]*url\s*:\s*['\"]([^'\"]+)['\"]",
        r"(?i)(/[a-z0-9_\-/]+\.(?:xlsx|xls|csv|json|do|action|ashx))",
        r"(?i)(/api/[a-z0-9_\-/]+)",
    ]
    found = set()
    for p in patterns:
        for m in re.findall(p, html):
            u = m if isinstance(m, str) else m[0]
            if any(k in u.lower() for k in ("fund", "etf", "holding", "asset", "pcf", "excel", "xls", "download", "api", "component", "constituent", "weight", "00406", "406a")):
                found.add(u)
    print(f"[FOUND {len(found)} candidate URLs]")
    for u in sorted(found):
        print(f"  {u}")
    # 抓 <script src=...>
    scripts = re.findall(r"<script[^>]+src=['\"]([^'\"]+)['\"]", html, re.I)
    print(f"[SCRIPTS {len(scripts)}]")
    for s in scripts[:20]:
        print(f"  {s}")
    # dump 前 3000 字節看 SPA 骨架
    print("---HTML head 3000---")
    print(html[:3000])
    print("---HTML tail 2000---")
    print(html[-2000:])
    return 0

if __name__ == "__main__":
    sys.exit(main())
