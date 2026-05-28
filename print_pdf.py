"""
print_pdf.py -- HTML to PDF via Playwright (headless Chromium)

Usage:
    python3 print_pdf.py input.html output.pdf

Playwright downloads its own Chromium on first install:
    pip install playwright
    python3 -m playwright install chromium
"""

import sys
import pathlib
import asyncio
from playwright.async_api import async_playwright


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

    async with async_playwright() as pw:
        browser = await pw.chromium.launch()
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
