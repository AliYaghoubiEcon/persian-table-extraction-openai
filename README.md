# Persian Table Digitization using OpenAI Vision

This project digitizes **Persian (Farsi) tables from images** using OpenAI’s multimodal models and converts them into clean, structured Excel files.

It is designed for tables with complex layouts, multi-line cells, and Persian text that are difficult to extract using traditional OCR tools.

---

## ✨ Features

- Extracts **entire table structure** from images (PNG / JPG)
- Preserves:
  - Original Persian text
  - Column order (left-to-right)
  - Multi-line cells (especially "توضیحات")
- Automatically detects table metadata from the title:
  - Province → `استان`
  - University name → `نام دانشگاه`
- Outputs:
  - One Excel file per image
  - One merged Excel containing all tables
  - One Excel with **one sheet per image**
- Resume-safe processing (progress saved to disk)
- Error logging for failed images

---

## 📂 Project Structure
project/
│
├── ensani_image/ # Input images (tables)
│ ├── image1.jpg
│ ├── image2.png
│
├── ensani_image/output_excels/
│ ├── per_image/ # One Excel per image
│ ├── merged_all.xlsx # All tables merged
│ ├── all_images_sheets.xlsx # One sheet per image
│ ├── progress.json # Resume state
│ └── errors.csv # Error log
│
├── digitize_tables.py # Main script
└── README.md


---

## ⚙️ Requirements

- Python 3.9+
- OpenAI Python SDK
- pandas
- tqdm
- openpyxl

Install dependencies:

```bash
pip install pandas tqdm openpyxl openai

  


