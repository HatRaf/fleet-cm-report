"""
print_pdf.py -- HTML to PDF via Playwright (headless Chromium)

Usage:
    python3 print_pdf.py input.html output.pdf

Playwright downloads its own Chromium on first install:
    pip install playwright
    python3 -m playwright install chromium
"""

import os
import sys
import pathlib
import asyncio
from playwright.async_api import async_playwright

# Optional system browser. Set CHROME_PATH when Playwright's bundled Chromium
# can't be downloaded (e.g. a locked-down cloud sandbox with a pre-installed
# browser). Empty -> use Playwright's own managed Chromium (normal local case).
CHROME_PATH = os.environ.get("CHROME_PATH", "").strip()


async def main():
    if len(sys.argv) != 3:
        print("Usage: python3 print_pdf.py input.html output.pdf")
        sys.exit(1)

    input_path  = pathlib.Path(sys.argv[1]).resolve()
    output_path = pathlib.Path(sys.argv[2]).resolve()

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        sys.exit(1)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    url = input_path.as_uri()

    launch_kwargs = {}
    if CHROME_PATH:
        launch_kwargs["executable_path"] = CHROME_PATH
        launch_kwargs["args"] = ["--no-sandbox", "--disable-setuid-sandbox"]

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(**launch_kwargs)
        page    = await browser.new_page()
        await page.goto(url, wait_until="networkidle")
        await page.wait_for_timeout(1500)   # allow fonts / deck-stage JS to settle
        await page.pdf(
            path=str(output_path),
            format="A4",
            print_background=True,
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
        )
        await browser.close()

    print(f"PDF written -> {output_path}")


asyncio.run(main())
