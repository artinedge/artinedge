# Digikala Comment Exporter

This repository contains a small Python utility that collects all public
comments for a product on [Digikala](https://www.digikala.com/) and stores them
in an Excel spreadsheet.

## Requirements

The script relies only on Python's standard library. Python 3.9 or newer is
recommended because the Excel writer uses `zipfile` and modern typing features.

## Usage

1. **Identify the product.** Open the Digikala product page in your browser and
   copy either the full URL (for example
   `https://www.digikala.com/product/dkp-7068663`), the `dkp-<id>` fragment, or
   just the numeric product ID (`7068663`).
2. **Run the exporter.** Execute the script with the product reference:

   ```bash
   python src/digikala_comments.py "https://www.digikala.com/product/dkp-7068663"
   ```

   By default the tool creates an Excel file named `digikala_comments.xlsx` in
   the current directory. Use `--output`/`-o` to choose a different filename.
3. **Adjust request pacing (optional).** Digikala may throttle rapid requests.
   If you encounter HTTP 429 responses, rerun the command with a longer delay
   between pages, for example:

   ```bash
   python src/digikala_comments.py dkp-7068663 --delay 1.5 --output comments.xlsx
   ```

### Command line options

* `--output` / `-o`: Path to the Excel file to create (default:
  `digikala_comments.xlsx`).
* `--delay`: Delay in seconds between consecutive API requests (default `0.5`).

The resulting spreadsheet contains one row per comment with metadata such as the
author, title, rating, purchase status, like/dislike counts and the positive and
negative points highlighted by the reviewer.
