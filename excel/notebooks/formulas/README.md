###### excel-functions
![](https://github.com/is-leeroy-jenkins/Halo-Kitty-Adventures/blob/main/static/images/excel-functions.png)

# 📘 Excel Formulas and Functions Tutorial (Excel 365 / 2021)

An end-to-end reference guide and tutorial for **Excel formulas and functions** targeting **Excel 365 / 2021**.
This repo is designed as both:

* A **Jupyter notebook–style tutorial** for step-by-step learning.
* A **reference manual** you can keep open alongside your workbooks.

It covers fundamental formula mechanics, core function families, modern **dynamic arrays**, and **LAMBDA** patterns, all with Excel-ready examples.

<a href="https://colab.research.google.com/github/is-leeroy-jenkins/Halo-Kitty-Adventures/blob/main/excel/notebooks/formulas/excel-formulas.ipynb" target="_parent">
<img src="https://colab.research.google.com/assets/colab-badge.svg" alt="Open In Colab"/></a>

---

## 📂 Repository Structure

> Adjust filenames/paths here to match your repo layout.

* `Excel_Formulas_Reference.ipynb`
  Jupyter Notebook version of the tutorial, with Markdown explanations and Excel formulas in code cells.

* `docs/Excel_Formulas_Reference.md` (optional)
  Markdown version of the tutorial for quick reading and documentation systems.

* `samples/` (optional)
  Example Excel workbooks that mirror sections of the tutorial (fundamentals, lookup, dynamic arrays, etc.).

---

## 🎯 Goals and Audience

This tutorial is aimed at:

* Analysts who want to move beyond simple SUM and IF.
* Developers and data professionals who treat Excel as a modeling platform.
* Anyone standardizing on **Excel 365 / 2021** and modern formula features (dynamic arrays, LAMBDA, etc.).

After working through the material, you should be able to:

* Design **robust analytical models** with clean, auditable formulas.
* Use **lookup, statistical, financial, and text functions** confidently.
* Apply **dynamic array** and **LAMBDA** patterns to reduce manual steps and helper columns.
* Document and debug formulas effectively for long-term maintainability.

---

## ✅ Prerequisites

* Microsoft **Excel 365 / 2021** (for dynamic arrays and LAMBDA functions).
* Basic familiarity with:

  * Selecting cells and ranges.
  * Entering formulas starting with `=`.
  * Saving and opening workbooks.

If you are using the notebook:

* Python 3 environment (for Jupyter).
* JupyterLab or VS Code with the Jupyter extension installed.

---

## 🚀 Getting Started

### 1. Clone or download the repo

```bash
git clone https://github.com/your-org/your-excel-formulas-repo.git
cd your-excel-formulas-repo
```

Or download as a ZIP and extract.

### 2. Open the notebook (optional, but recommended)

Open `Excel_Formulas_Reference.ipynb` in:

* **JupyterLab** (`jupyter lab` from the repo directory), or
* **VS Code** with the Jupyter extension.

You can read the explanations in Markdown cells and copy the Excel formulas from code cells into your own workbook.

### 3. Open the Excel workbooks (if provided)

* Open any sample file in the `samples/` folder.
* Follow along with the formulas as you read the corresponding section in the notebook or markdown tutorial.

---

## 🧩 Tutorial Overview

The tutorial is organized into **function families**, each section following a consistent pattern:

1. **High-level explanation** of the purpose and typical use cases.
2. **Design guidance** (naming, volatility, readability).
3. **Excel-ready formula examples** that you can copy into worksheets.

### Sections at a Glance

* **Formula Fundamentals**

  * How Excel evaluates formulas (precedence, parentheses).
  * Relative vs. absolute vs. mixed references.
  * Named ranges, Tables, and `IFERROR` / `IFNA` for robustness.

* **Math & Trigonometry**

  * Summation and aggregation: `SUM`, `PRODUCT`, `SUMPRODUCT`.
  * Rounding and precision: `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`.
  * Modular arithmetic and scaling: `MOD`, `POWER`, `SQRT`, `LOG`, `LN`.
  * Trigonometry and angle conversions.

* **Logical**

  * Classic branching: `IF`, `IFS`.
  * Category mapping: `SWITCH`.
  * Boolean composition: `AND`, `OR`, `NOT`, `XOR`.
  * Error handling: `IFERROR`, `IFNA`.

* **Date & Time**

  * Date/time storage model and serial values.
  * Constructing dates and times: `DATE`, `TIME`, `DATEVALUE`.
  * Extracting components: `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`.
  * Calendar arithmetic: `EDATE`, `EOMONTH`, `WORKDAY`, `NETWORKDAYS`.
  * Durations and fractions: `DATEDIF`, `YEARFRAC`.

* **Text**

  * Cleaning and normalization: `TRIM`, `CLEAN`, `UPPER`, `LOWER`, `PROPER`.
  * Extraction: `LEFT`, `RIGHT`, `MID`.
  * Search and replace: `SEARCH`, `FIND`, `REPLACE`, `SUBSTITUTE`.
  * Concatenation: `CONCAT`, `TEXTJOIN`.
  * Formatting: `TEXT`.
  * Tokenization (365+): `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`.

* **Lookup & Reference**

  * Modern lookup: `XLOOKUP` and `XMATCH`.
  * Classic lookup: `VLOOKUP`, `HLOOKUP`.
  * Flexible retrieval: `INDEX` + `MATCH` (1D and 2D lookups).
  * Projection helpers: `CHOOSECOLS`, `CHOOSEROWS`.
  * Structural info: `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `ADDRESS`.
  * PivotTable integration: `GETPIVOTDATA`.

* **Statistical**

  * Descriptive stats: `AVERAGE`, `MEDIAN`, `MODE`, `MIN`, `MAX`.
  * Conditional stats: `AVERAGEIF(S)`, `COUNTIF(S)`, `MAXIFS`, `MINIFS`.
  * Distribution and shape: `STDEV.P/S`, `VAR.P/S`, `SKEW`, `KURT`.
  * Ranks and percentiles: `RANK.EQ`, `PERCENTILE.INC/EXC`, `QUARTILE`.
  * Correlation and simple regression: `CORREL`, `FORECAST.LINEAR`, `LINEST`.
  * Basic tests: `T.TEST`, `Z.TEST`, `CHISQ.TEST`, `F.TEST`.

* **Financial**

  * Time value of money: `PMT`, `FV`, `PV`, `RATE`, `NPER`.
  * Cash-flow analysis: `NPV`, `XNPV`, `IRR`, `XIRR`.
  * Depreciation: `SLN`, `DDB`, `DB`.
  * Loan schedule summaries: `CUMIPMT`, `CUMPRINC`.
  * Bonds and duration: `PRICE`, `YIELD`, `PRICEDISC`, `YIELDDISC`, `DURATION`, `MDURATION`.
  * Accrued interest and coupon functions: `ACCRINT`, `ACCRINTM`, `COUP*`.
  * Effective vs. nominal rates: `EFFECT`, `NOMINAL`.

* **Database & Information**

  * Database functions: `DSUM`, `DCOUNT`, `DAVERAGE`, `DMAX`, `DMIN`.
  * Metadata and cell info: `CELL`, `INFO`, `TYPE`.
  * Type and error checking: `ISNUMBER`, `ISTEXT`, `ISBLANK`, `ISERROR`, `ISNA`, `ERROR.TYPE`.
  * PivotTable retrieval with `GETPIVOTDATA`.

* **Dynamic Arrays and LAMBDA (Excel 365+)**

  * Spilling basics: `SEQUENCE`, `UNIQUE`, `FILTER`, `SORT`, `SORTBY`.
  * Reshaping: `CHOOSECOLS`, `CHOOSEROWS`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `WRAPROWS`, `WRAPCOLS`, `EXPAND`, `TOCOL`, `TOROW`.
  * Higher-order operations: `MAP`, `BYROW`, `BYCOL`, `REDUCE`, `SCAN`.
  * Local variables and inline composition: `LET`.
  * Custom functions in pure Excel: `LAMBDA`, `MAKEARRAY`.

* **Auditing & Optimization**

  * Formula auditing tools (Trace Precedents, Dependents, Evaluate Formula, Watch Window).
  * Guarding interfaces with `IFERROR` / `IFNA`.
  * Using Tables and Named Ranges for stability.
  * Reducing volatility and repeated computation via `LET`.
  * Documentation patterns for production spreadsheets.

---

## 🧪 How to Use This Tutorial

* Use the notebook or markdown file as a **reference**:

  * Keep it open while building or refactoring workbooks.
  * Copy formulas directly into cells and adapt ranges/names.

* Treat each section as a **mini-module**:

  * Start at “Formula Fundamentals” if you are reinforcing basics.
  * Jump directly to “Lookup & Reference”, “Dynamic Arrays”, or “Financial” when working on those problems.

* Integrate into your workflow:

  * Save the notebook or markdown in your personal toolkit repo.
  * Build small test sheets for each function family (mirroring the examples).
  * Gradually refactor old workbooks toward the patterns shown here.

---

## 🧱 Design Principles Used in the Tutorial

* Emphasis on **clarity over cleverness**:

  * Prefer readable formulas with helper cells to overly dense one-liners.
* Focus on **robustness and auditability**:

  * Guard error-prone steps.
  * Use naming and documentation patterns that survive hand-offs.
* Use of **modern Excel features**:

  * Dynamic arrays instead of array-entered formulas.
  * LAMBDA and LET for reusable logic and reduced duplication.

---

## 🤝 Contributing

Suggestions and improvements are welcome, especially for:

* Additional practical examples (e.g., budgeting, forecasting, dashboards).
* Edge cases where specific functions behave unexpectedly.
* More sample workbooks illustrating advanced dynamic array or LAMBDA patterns.

You can propose changes via:

1. Forking the repo.
2. Creating a branch for your enhancements.
3. Opening a pull request with a clear description of the changes.

---

## 📄 License

Specify your license choice here (for example):

This project is licensed under the **MIT License**.
See the `LICENSE` file for details.

---

## 📎 Quick Start Copy-Paste

If you only need the main resource quickly:

* Primary tutorial notebook: `Excel_Formulas_Reference.ipynb`

Open it, scroll section by section, and start dropping formulas directly into your own Excel models.
