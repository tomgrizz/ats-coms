"""
Make the one-pager HTML fully self-contained:
  - Embed local images as base64 data URIs
  - Download Google Fonts and inline them as base64
Output: a single .html file that renders correctly with no internet or local assets needed.
"""
import base64, os, re, urllib.request, time

BASE     = r"C:\Users\lomu-\Documents\GitHub\lake-ontario-atlantic-salmon-communication-plan\project\Lake Ontario Atlantic Salmon Communication Plan"
HTML_IN  = os.path.join(BASE, "Bring Back the Salmon One-Pager.html")
HTML_OUT = os.path.join(BASE, "Bring Back the Salmon - Program Brief (Self-Contained).html")

with open(HTML_IN, "r", encoding="utf-8") as f:
    html = f.read()

# ── 1. Embed local images ─────────────────────────────────────────────────────
MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".svg": "image/svg+xml", ".gif": "image/gif", ".webp": "image/webp"}

def to_data_uri(rel_path):
    full = os.path.join(BASE, rel_path.replace("/", os.sep))
    if not os.path.exists(full):
        print(f"  [skip] not found: {full}")
        return None
    ext  = os.path.splitext(rel_path)[1].lower()
    mime = MIME.get(ext, "application/octet-stream")
    b64  = base64.b64encode(open(full, "rb").read()).decode("ascii")
    return f"data:{mime};base64,{b64}"

def embed_local_images(html):
    def replacer(m):
        attr, path = m.group(1), m.group(2)
        uri = to_data_uri(path)
        return f'{attr}="{uri}"' if uri else m.group(0)
    # Match src="assets/..." or src='assets/...'
    return re.sub(r'(src)=["\']([^"\']*assets/[^"\']+)["\']', replacer, html)

print("Embedding local images...")
html = embed_local_images(html)

# ── 2. Download & embed Google Fonts ─────────────────────────────────────────
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"

# Find the Google Fonts CSS URL in the HTML
font_link_match = re.search(
    r'<link\s[^>]*href="(https://fonts\.googleapis\.com/css2[^"]+)"[^>]*>',
    html
)

if font_link_match:
    fonts_url = font_link_match.group(1)
    print(f"Fetching Google Fonts CSS...")
    req = urllib.request.Request(fonts_url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req) as resp:
        fonts_css = resp.read().decode("utf-8")

    # Download every font file URL in the CSS and replace with base64
    font_urls = re.findall(r'url\((https://fonts\.gstatic\.com/[^)]+)\)', fonts_css)
    print(f"  Embedding {len(font_urls)} font file(s)...")
    font_cache = {}
    for url in font_urls:
        if url in font_cache:
            continue
        r = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(r) as resp:
            data = resp.read()
        fmt  = "woff2" if ".woff2" in url else "woff" if ".woff" in url else "truetype"
        b64  = base64.b64encode(data).decode("ascii")
        font_cache[url] = f"url(data:font/{fmt};base64,{b64}) format('{fmt}')"
        time.sleep(0.05)  # be polite to Google's servers

    def font_replacer(m):
        return font_cache.get(m.group(1), m.group(0))

    embedded_fonts_css = re.sub(
        r'url\((https://fonts\.gstatic\.com/[^)]+)\)',
        font_replacer,
        fonts_css,
    )

    # Remove all Google Fonts <link> tags (preconnect + stylesheet)
    html = re.sub(r'\s*<link\s[^>]*(fonts\.googleapis\.com|fonts\.gstatic\.com)[^>]*/?\s*>', '', html)

    # Inject embedded fonts as <style> before </head>
    font_block = f"<style>\n/* Google Fonts — embedded */\n{embedded_fonts_css}\n</style>"
    html = html.replace("</head>", font_block + "\n</head>", 1)
    print("  Google Fonts embedded.")
else:
    print("  No Google Fonts link found — skipping.")

# ── 3. Save ───────────────────────────────────────────────────────────────────
with open(HTML_OUT, "w", encoding="utf-8") as f:
    f.write(html)

kb = round(os.path.getsize(HTML_OUT) / 1024, 1)
print(f"\nDone — self-contained HTML saved ({kb} KB):")
print(f"  {HTML_OUT}")
