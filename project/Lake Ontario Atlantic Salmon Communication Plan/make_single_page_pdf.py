"""
Render the one-pager HTML to a PDF (letter size, as many pages as needed).
"""
import subprocess, os, tempfile
from pypdf import PdfReader

BASE    = r"C:\Users\lomu-\Downloads\Lake Ontario Atlantic Salmon Communication Plan"
HTML_IN = os.path.join(BASE, "Bring Back the Salmon One-Pager.html")
PDF_OUT = os.path.join(BASE, "Bring Back the Salmon - Program Brief.pdf")
CHROME  = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

with open(HTML_IN, "r", encoding="utf-8") as f:
    html = f.read()

inject = """<style id="__print__">
@media print {
  @page { size: 8.5in 11in; margin: 0; }
  html, body { background: #fff !important; }
  .stage { padding: 0 !important; display: block !important; }
  .page  { box-shadow: none !important; width: 8.5in !important; }
}
</style>"""
modified = html.replace("</head>", inject + "\n</head>", 1)

tmp = tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8")
tmp.write(modified)
tmp.close()

try:
    file_url = "file:///" + tmp.name.replace("\\", "/").replace(" ", "%20")
    print("Rendering HTML with Chrome headless...")
    subprocess.run([
        CHROME,
        "--headless=new", "--disable-gpu", "--no-sandbox",
        "--run-all-compositor-stages-before-draw",
        "--allow-file-access-from-files",
        f"--print-to-pdf={PDF_OUT}",
        "--print-to-pdf-no-header",
        "--no-pdf-header-footer",
        file_url,
    ], check=True, capture_output=True, timeout=45)
finally:
    os.unlink(tmp.name)

reader = PdfReader(PDF_OUT)
n = len(reader.pages)
kb = round(os.path.getsize(PDF_OUT) / 1024, 1)
print(f"Done - PDF saved ({n} page(s), {kb} KB):")
print(f"  {PDF_OUT}")
