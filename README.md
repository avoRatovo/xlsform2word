# 📦 xlsform2word

`xlsform2word` is an R package that converts **XLSForm questionnaires** (from KoboToolbox or ODK) into well-formatted **Word documents (.docx)**.

It helps researchers, monitoring officers, evaluators, and humanitarian staff to document and review surveys in a clean, human-readable format — especially for sharing with non-technical stakeholders.

---

## 🚀 Key Features

- Automatically reads `.xls` or `.xlsx` XLSForms
- Detects **sections**, **nested groups**, and **repeat sections**
- Supports both `label::French` and `label::English`
- Smart formatting in Word, including:
  - Section headers with numbering
  - Question and answer layout
  - Select_one / select_multiple options displayed properly
  - `calculate` fields shown in red
  - Notes in italic
  - Display conditions (`relevant`) written clearly in plain text
- Handles deep group nesting (up to 3+ levels)

---

## 📦 Installation

You can install the package directly from GitHub:

```r
install.packages("devtools")  # if not already installed

devtools::install_github("avoRatovo/xlsform2word")
