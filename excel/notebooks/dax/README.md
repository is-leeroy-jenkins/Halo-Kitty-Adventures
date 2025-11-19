
# 📘 Comprehensive DAX Tutorial

A practical, Power BI–ready, icon-rich guide to mastering Data Analysis Expressions
for Power BI, Analysis Services, and Excel Power Pivot.

---

## 🐱 Overview

This repository is a one-stop DAX learning suite for analysts, developers, and anyone who wants to build robust, context-aware Power BI reports.

* 📖 Jupyter Notebook tutorial (.ipynb) — copy-paste and run!
* 🧩 Sectional notebooks for deep-diving or quick reference
* 📊 Power BI visuals with real-world DAX measures
* 🎯 DAIM/DoD Data Analytics learning progression

---

## 📁 Contents

| File                        | Description                                     |
| --------------------------- | ----------------------------------------------- |
| DAX_Tutorial_Complete.ipynb | ⭐ Start here: Full step-by-step DAX walkthrough |
| DAX_Chunk1.ipynb            | Basics, syntax, aggregation, CALCULATE          |
| DAX_Chunk2.ipynb            | Filters, time, relationships, logic/text        |
| DAX_Chunk3.ipynb            | Table/row, ranking, hierarchy, VAR patterns     |
| DAX_Chunk4.ipynb            | Power BI visual DAX examples                    |
| DAX_Chunk5.ipynb            | Title, appendix, best practices, debug          |
| assets/                     | (Optional) Images, diagrams, PBIX files         |
| README.md                   | This file                                       |

---

## 🏗️ Learning Flow

1. **Start with** DAX_Tutorial_Complete.ipynb
2. Work section-by-section or jump to a topic using the table of contents
3. Try out code cells in your own Power BI models
4. Reference chunked notebooks for focused study
5. Review visual scenarios for real business reporting use

---

## 🚦 What You’ll Learn

* How DAX computes with **row and filter context**
* Aggregation, iterators (SUMX, AVERAGEX), and context transition
* Building robust time intelligence (YTD, QTD, MAT, YoY, etc.)
* Using CALCULATE, ALL, ALLEXCEPT, and KEEPFILTERS
* Relationship functions: RELATED, USERELATIONSHIP, CROSSFILTER
* Hierarchies with PATH/PATHITEM
* Ranking and windowing (RANKX, OFFSET, INDEX)
* Real-world visuals: top N, KPI cards, funnel analysis
* Debugging, performance tips, and best practices

---

## 📊 Power BI Visual Examples

Each visual in the notebook series comes with:

* Visual description, key DAX logic, and best use case
* Ready-to-use measures for:

  * 📦 Matrix (Sales, Profit, Margin)
  * 📈 Line chart (YoY, MoM trends)
  * 🏅 Top N ranking bar charts
  * 🎯 KPI cards (% to goal)
  * 🔻 Funnels (Orders → Delivered → Paid)

---

## 🧮 DAX Function Groups

| Icon | Area                   | Key Functions / Concepts                  |
| ---- | ---------------------- | ----------------------------------------- |
| 📐   | Syntax/Eval Model      | Context, CALCULATE, filter transition     |
| ➕    | Aggregation            | SUM, AVERAGE, MIN, MAX, COUNTROWS         |
| 🔁   | Iterators              | SUMX, AVERAGEX, COUNTX, FILTERX           |
| 🎯   | Filter Manipulation    | CALCULATE, ALL, ALLEXCEPT, FILTER         |
| 📊   | Time Intelligence      | YTD, QTD, MTD, SAMEPERIODLASTYEAR         |
| 🔄   | Relationships          | RELATED, USERELATIONSHIP, CROSSFILTER     |
| 🧱   | Table/Row Constructors | ADDCOLUMNS, ROW, SELECTCOLUMNS, SUMMARIZE |
| 🔍   | Ranking/Windows        | RANKX, OFFSET, INDEX, WINDOW              |
| 🧩   | Hierarchies            | PATH, PATHITEM, PATHCONTAINS              |
| 🧮   | Text/Logic/Math        | CONCATENATEX, DIVIDE, SWITCH, IF          |
| 🧠   | Debug/Best Practices   | ALLSELECTED, DAX Studio, VAR, safe DIVIDE |

---

## 🛠️ Getting Started

**Requirements:**

* Python 3.x
* Jupyter Notebook (or VS Code Jupyter)
* Power BI Desktop (for PBIX testing)

**Install Jupyter:**

```
pip install notebook
```

**Run the notebook:**

```
jupyter notebook
```

Then open `DAX_Tutorial_Complete.ipynb` from your browser.

---

## 🤝 Contributing

* Fork and PR to add advanced DAX scenarios, visuals, or lessons
* Add PBIX reports, screenshots, or appendix materials to `assets/`
* Raise issues for clarification or bug reports

---

## 📜 License

MIT License — free for training, DAIM/DoD modernization, or internal analytics upskilling.

---

## 🙋 Questions?

Open an issue or reach out to the repo author.
Happy DAXing!
