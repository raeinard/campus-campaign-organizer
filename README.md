# Campus Campaign Organizer

**Clean messy Google Form responses → Group students by School → Residence → Block → Floor → Generate per‑school Word documents for door‑to‑door campaigning.**

## Problem

Students fill a Google Form with fields: Full Names, WhatsApp Number, School, Res, Room no. & Block, availability.

Challenges:
- Residence names are inconsistent: `"2k beds res 3"`, `"2000beds res3"`, `"res3"` → all should become `"Res 3"`.
- Unknown residences (e.g., `"Arebeng"`) must be preserved.
- `Room no. & Block` is messy: `"E101"` (non‑2000 beds) means Block E, Floor 1, Room 101.  
  Free‑text like `"2000 beds res 3, Block B, 2nd floor, Unit 12"` must be parsed.
- Need to group students so your team visits **one floor of one block** in one residence for a specific school.

## Solution

This Python pipeline (runs in **Google Colab**) does:

1. **Standardises residence names** using regex + fuzzy matching.
2. **Parses `Room no. & Block`** into `floor`, `block`, `room_number`, `unit` (special rule for non‑2000 beds: `E101` → Block E, Floor 1, Room 101).
3. **Groups** by `School → Residence → Block → Floor`.
4. **Generates separate Word documents** per school (Name, WhatsApp, Room, Unit, Availability).
5. **Zips all documents** for download.

## Input format (CSV/Excel from Google Forms)

Expected column names:
- `Full Names`
- `WhatsApp Number`
- `School`
- `Res`
- `Room no. & Block`
- `When are you available`
- `Time available` (optional)

## How to use
# Online Google Colab version
1. Click the **Open in Colab** button below.
2. Upload your CSV/Excel file when prompted.
3. Run all cells – download the zip file.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/raeinard/campus-campaign-organizer/blob/main/campaign_organizer.ipynb)

# Local Python version

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/campus-campaign-organizer.git
   cd campus-campaign-organizer



## Output

- `campaign_plans_by_school.zip` – one Word document per school.
- `cleaned_responses.csv` – cleaned dataset.

## Customisation

- Add new residence variations → edit `manual_map` in the notebook.
- Change the special pattern for non‑2000 beds → edit regex `r'^([A-Z])(\d)(\d{2,3})$'`.

## Files

- `campaign_organizer.ipynb` – main Colab notebook.
- `requirements.txt` – Python dependencies.
- `data/sample_responses.csv` – anonymised sample data.
- `.gitignore` – ignores generated files.


