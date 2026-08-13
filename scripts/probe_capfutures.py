"""測 capfutures /api/holdings 各種參數."""
import json, requests

BASE = "https://etf.capfutures.com"
HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-TW,zh;q=0.9",
    "Referer": f"{BASE}/?view=holdings&fund=00406A",
    "Origin": BASE,
}

def hit(url):
    try:
        r = requests.get(url, headers=HDR, timeout=30)
        print(f"\n[GET] {url}\n  status={r.status_code} len={len(r.content)} ct={r.headers.get('Content-Type','?')}")
        body = r.text[:2500]
        print(f"  body: {body}")
        if r.headers.get("Content-Type", "").startswith("application/json"):
            try:
                j = r.json()
                if isinstance(j, dict):
                    print(f"  keys: {list(j.keys())[:20]}")
                    if "data" in j:
                        d = j["data"]
                        if isinstance(d, list):
                            print(f"  data list len={len(d)}")
                            if d: print(f"  data[0]: {json.dumps(d[0], ensure_ascii=False)[:500]}")
                        elif isinstance(d, dict):
                            print(f"  data keys: {list(d.keys())[:20]}")
                elif isinstance(j, list):
                    print(f"  list len={len(j)}")
                    if j: print(f"  [0]: {json.dumps(j[0], ensure_ascii=False)[:500]}")
            except Exception as e:
                print(f"  json parse fail: {e}")
    except Exception as e:
        print(f"[FAIL] {url}: {e}")

for endpoint in [
    "/api/holdings",
    "/api/holdings?fund=00406A",
    "/api/holdings?fundCode=00406A",
    "/api/holdings?code=00406A",
    "/api/holdings/00406A",
    "/api/history",
    "/api/history?fund=00406A",
    "/api/me",
]:
    hit(BASE + endpoint)
