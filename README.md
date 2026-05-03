# IT3040 Assignment 1 – Singlish Transliteration Testing

**Student:** IT23763258
**Target App:** https://www.pixelssuite.com/chat-translator
**Tool:** Playwright (Python)
**GitHub Repository:** https://github.com/sanduni-hettiarachchi/IT23763258_ITPM_Assignment1

---

## Project Structure

---

## Prerequisites

- **Python 3.11 or 3.12** — [Download](https://www.python.org/downloads/)
- **Google Chrome** (recommended) — [Download](https://www.google.com/chrome/)
- **VS Code** (recommended editor) — [Download](https://code.visualstudio.com/)

> **Windows users:** Make sure Python is added to your PATH during installation.

---

## Setup & Installation (One-Time)

Open **Command Prompt** or **VS Code Terminal**, then run:

```bash
# Step 1 – Navigate into the project folder
cd path\to\IT23763258

# Step 2 – Upgrade pip
pip install -U pip

# Step 3 – Install required Python packages
pip install playwright openpyxl

# Step 4 – Install Playwright browsers
playwright install
```

---

## How to Run the Tests

From inside the project folder in your VS Code terminal, run:

```bash
python IT23763258_test_automation.py --excel "IT23763258_Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 8000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

---

## What Happens When You Run It

1. **Browser opens** and loads `https://www.pixelssuite.com/chat-translator`
2. **Each row** in the Excel file is read in order
3. The Singlish input is **typed** into the Chat Sinhala input box
4. The **Transliterate** button is clicked
5. The actual Sinhala output is **captured** and written back to Excel
6. Status is set to **PASS** (actual == expected) or **FAIL** (mismatch)
7. Results are **saved** to the same Excel file after each row

---

## Checking Results

After the run completes:

1. Open `IT23763258_Assignment 1 - Test cases.xlsx`
2. Check columns **E (Actual Output)** and **F (Status)**
3. All 50 negative test cases should show **FAIL** (they were designed to fail)
4. All 10 positive test cases should show **PASS**

---

## Test Case Summary

| Type | Count |
|------|-------|
| Positive (expected Pass) | 10 |
| Negative (expected Fail) | 50 |
| **Total** | **60** |

### Singlish Input Types Covered (all 24 required)

1. Question forms
2. Command forms
3. Greetings
4. Requests
5. Responses
6. Repeated Words
7. Inputs with Punctuation Marks
8. Romanization / Spelling Variants
9. Isolated English Word Insertions in Singlish
10. Multi-Word English Phrases in Singlish
11. English Digital Terms in Singlish
12. Platform/App Names in Singlish
13. English Abbreviations/Acronyms in Singlish
14. English Clipped Forms in Singlish
15. Place Names Embedded in Singlish
16. Person Names Embedded in Singlish
17. Inputs with Numbers and Numeric Suffixes
18. Inputs with Currency
19. Inputs with Time Formats
20. Inputs with Dates
21. Inputs with Unit of Measurements
22. Inputs with Slang and Casual Phrasing
23. Online Identifiers in Singlish
24. Inputs Containing Emojis

---

## Troubleshooting

**"playwright: command not found"**
→ Try: `python -m playwright install`

**Browser opens but no typing happens**
→ Increase `--wait-ms` to 10000 and `--slow-mo-ms` to 500

**Excel not found error**
→ Make sure `IT23763258_Assignment 1 - Test cases.xlsx` is in the same folder as `IT23763258_test_automation.py`

**Output column stays empty**
→ The site may have changed its layout. Try running without `--headless` to watch what happens.

**Unicode/encoding errors on Windows**
→ Run: `chcp 65001` in Command Prompt before running the script