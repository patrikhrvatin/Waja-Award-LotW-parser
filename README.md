# WAJA (Worked All Japan Award) Log Analyzer - by 9A5CW and a help of AI

A Python tool designed to parse ADIF log files (such as LoTW reports), analyze amateur radio contacts with Japan, track progress for the **WAJA (Worked All Japan)** award, and generate detailed reports and formatted Excel spreadsheets.

---

## Features & What the Script Does

* **ADIF Parsing:** Scans large `.adi` log files (like `lotwreport.adi`) to extract all confirmed QSOs with Japanese stations holding valid JARL prefecture codes.
* **Award Progress Tracking:**
* Checks **Basic WAJA** status (all 47 prefectures).
* Tracks multi-band progress (e.g., 5-band and 9-band WAJA completion).


* **Band-by-Band Statistics:** Calculates confirmed prefecture counts and percentages across standard amateur radio bands (from 160M down to 6M).
* **Smart Band Leads Analysis:** Cross-references your logs to find "multiband stations"—stations you've already worked on other bands, but are missing on a specific band where you still need their prefecture (great for scheduling!).
* **Report Generation:** Automatically outputs clean text files and spreadsheets containing:
* Official JARL application list (`waja_jarl_application.txt`)
* Missing basic WAJA prefectures (`missing_basic_waja.txt`)
* Missing prefectures broken down by band (`missing_by_band.txt`)
* Top active stations list (`top_active_stations.txt`)
* Potential schedule leads for missing band prefectures (`missing_band_leads.txt`)
* Formatted Excel sheets with color coding (`waja_lotw_report.xlsx`, `missing_prefectures.xlsx`)



---

## Requirements

The script requires Python 3 and the following external libraries:

* `pandas`
* `openpyxl`

---

## Installation & Setup

1. Make sure you have Python installed.
2. Install the required dependencies via terminal:
```bash
pip install -r requirements.txt

```


3. Place your ADIF file (named `lotwreport.adi`) in the same folder as the script.
4. Run the script:
```bash
python waja11en.py

```
