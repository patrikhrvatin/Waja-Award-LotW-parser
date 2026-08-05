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
## 🚀 Example Output

Here is what running the script looks like in practice when scanning a large ADIF log (over 105,000 records):

```text
C:\Users\9A5CW\Desktop\waja>waja11en.py
[*] Starting scan of file 'lotwreport.adi'...
[i] Read 25000 records...
[i] Read 50000 records...
[i] Read 75000 records...
[i] Read 100000 records...

[+] Analysis complete! Total ADIF records checked: 105103
[+] Found 1737 confirmed Japan QSOs with a valid JARL prefecture.
[+] Službena JARL prijavna lista stvorena u: 'waja_jarl_application.txt'
[+] Missing basic WAJA list saved to: 'missing_basic_waja.txt'
[+] Missing by band report saved to: 'missing_by_band.txt'
[+] Top active stations report saved to: 'top_active_stations.txt'
[+] Missing band leads report saved to: 'missing_band_leads.txt'

================ WAJA BAND STATISTICS ================
-> Band 160M : Confirmed 22/47 Prefectures (46.8%)
-> Band 80M  : Confirmed 18/47 Prefectures (38.3%)
-> Band 40M  : Confirmed 43/47 Prefectures (91.5%)
-> Band 30M  : Confirmed 42/47 Prefectures (89.4%)
-> Band 20M  : Confirmed 37/47 Prefectures (78.7%)
-> Band 17M  : Confirmed 44/47 Prefectures (93.6%)
-> Band 15M  : Confirmed 38/47 Prefectures (80.9%)
-> Band 12M  : Confirmed 43/47 Prefectures (91.5%)
-> Band 10M  : Confirmed 35/47 Prefectures (74.5%)
-> Band 6M   : Confirmed 41/47 Prefectures (87.2%)

================ AWARD TRACKER SUMMARY ================
-> Basic WAJA Status  : Worked 47/47 unique Prefectures.
-> 5-Band WAJA Status : 43 Prefectures completed on 5+ bands.
   Completed: 01 (Hokkaido), 02 (Aomori), 03 (Iwate), 04 (Akita), 05 (Yamagata)
...
-> 9-Band WAJA Status : 22 Prefectures completed on 9+ bands.
   Completed: 01 (Hokkaido), 03 (Iwate), 06 (Miyagi), 08 (Niigata), 10 (Tokyo),
11 (Kanagawa), 12 (Chiba), 13 (Saitama), 14 (Ibaraki), 15 (Tochigi), 17 (Yamanashi), 18 (Shizuoka), 20 (Aichi), 25 (Osaka), 27 (Hyogo), 28 (Toyama), 31 (Okayama), 32 (Shimane), 38 (Ehime), 40 (Fukuoka), 41 (Saga), 43 (Kumamoto)
-> Congratulations! You have worked all 47 prefectures at least once!
-> Missing by band breakdown saved to 'missing_by_band.txt'
-> Top active stations report saved to 'top_active_stations.txt'
======================================================
