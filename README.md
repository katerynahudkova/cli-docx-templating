# DOCX Templater CLI
This project provides a small CLI to generate customized Word documents from a template and data stored in a CSV file.

The tool replaces placeholders of the form {var} inside your Word template with values from matching CSV columns.

Python-docx was used for document processing and template filling. Pandas was used for its useful handling with CSV files.

## Features
- **Placeholder replacement**  
  Supports `{var}` placeholders inside:
  - paragraphs,
  - tables,
  - headers and footers. 

- **Validation**  
  Warns if placeholders in the template are missing in the CSV, or if the CSV has unused columns.

- **Batch generation**  
  Creates one output `.docx` per row in the CSV.

- **Custom filenames**  
  Output files are named using a user-defined pattern, e.g.:  
  NDA_{side1_name_en}_{year}.docx

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/<your-github-id>/cli-docx-templating.git
cd cli-docx-templating
pip install -r requirements.txt
```

## Usage
```bash
python -m src.docx_templater.cli \
  --template samples/template.docx \
  --csv samples/values.csv \
  --outdir output \
  --pattern "NDA_{side1_name_en}_{year}.docx"
```

Arguments

--template, -t → Path to the Word template (.docx)

--csv, -c → Path to the CSV file with data

--outdir, -o → Output directory where documents will be written

--pattern, -p → Filename pattern using column names in {var} style

## Example

Suppose your template.docx contains placeholders:

```
This NDA is between {side1_name_en} and {side2_name_en}, signed in {city_ua}, {year}.
```

And your values.csv has:

```
side1_name_en	side2_name_en	city_ua	year
Alice	Bob	Київ	2024
Company	Partners	Львів	2025
```

Running:

```bash
python -m src.docx_templater.cli \
  --template samples/template.docx \
  --csv samples/values.csv \
  --outdir output \
  --pattern "NDA_{side1_name_en}_{year}.docx"
```

Will produce:
```
output/
├── NDA_Alice Ltd_2024.docx
└── NDA_X Company_2025.docx
```

## Known Limitations

- Only {var} placeholder style is supported.

- Does not handle advanced Word features (e.g., content controls, images).

- Filenames must be valid for your operating system (invalid characters will cause save errors).

- CSV must be UTF-8 encoded.
