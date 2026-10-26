<div align="center">

# 🌿 Ferns & Petals — Sales Analysis Dashboard

**An end-to-end Microsoft Excel project** that prepares three datasets, models them into a single analytical table, and turns 1,000 orders from 100 customers into an interactive dashboard covering sales trends, customer behaviour, product performance and delivery speed.

</div>

---

## 📑 Table of Contents

- [Overview](#-overview)
- [Key Results](#-key-results)
- [Repository Structure](#-repository-structure)
- [Datasets](#-datasets)
- [Data Preparation Summary](#-data-preparation-summary)
- [What Was Done in Excel](#-what-was-done-in-excel)
- [Dashboard Overview](#-dashboard-overview)
  - [1. All Occasions Dashboard](#1-All-Occasions-Dashboard)
  - [2. Anniversary Dashboard](#2-Anniversary-Dashboard)
  - [3. Birthday Dashboard](#3-Birthday-Dashboard)
  - [4. Diwali Dashboard](#4-Diwali-Dashboard)
  - [5. Holi Dashboard](#5-Holi-Dashboard)
  - [6. Raksha Bandhan Dashboard](#6-Raksha-Bandhan-Dashboard)
  - [7. Valentine's Day Dashboard](#7-Valentine's-Day-Dashboard)

- [Analysis & Visualizations](#-analysis--visualizations)
  - [1. Total Revenue](#1-total-revenue)
  - [2. Average Order & Delivery Time](#2-average-order--delivery-time)
  - [3. Monthly Sales Performance](#3-monthly-sales-performance)
  - [4. Top Products by Revenue](#4-top-products-by-revenue)
  - [5. Customer Spending Analysis](#5-customer-spending-analysis)
  - [6. Sales Performance of Top 5 Products](#6-sales-performance-of-top-5-products)
  - [7. Top 10 Cities by Number of Orders](#7-top-10-cities-by-number-of-orders)
  - [8. Order Quantity vs Delivery Time](#8-order-quantity-vs-delivery-time)
  - [9. Revenue Comparison Between Occasions](#9-revenue-comparison-between-occasions)
  - [10. Product Popularity by Occasion](#10-product-popularity-by-occasion)
- [Category Performance](#-category-performance)
- [Business Recommendations](#-business-recommendations)
- [Known Data Notes](#-known-data-notes)
- [Tech Stack](#-tech-stack)
- [How to Use](#-how-to-use)
- [Documentation](#-documentation)
- [Author & Contact](#-author--contact)

---

## 🧭 Overview

Ferns & Petals (FNP) is a gifting company that delivers flowers, cakes, plants, soft toys, sweets and mugs for occasions such as Diwali, Raksha Bandhan, Holi, Valentine's Day, Birthdays and Anniversaries. This project analyses its transactional data to answer one core question:

> **When, what, where and for whom does FNP sell — and how quickly does it deliver?**

The workflow covers the full analytics lifecycle:

```
Raw CSVs → Data Quality Checks → Excel Tables → Derived Columns (Revenue, Delivery Days…)
→ PivotTables → PivotCharts → Interactive Dashboard (Slicer + Timelines)
→ Insights → Business Recommendations
```

All analysis was performed in `Ferns_Petals.xlsx` using **Excel Tables, PivotTables, PivotCharts, a Slicer, Timelines and formulas** (including `CORREL`).

---

## 🎯 Key Results

| Metric | Value |
|---|---|
| **Total Revenue** | ₹35,20,984 |
| **Total Orders** | 1,000 |
| **Total Customers** | 100 |
| **Average Order Value** | ₹3,521 |
| **Average Spend per Customer** | ₹35,210 |
| **Average Order → Delivery Time** | 5.53 days (min 1, max 10) |
| **Peak Sales Month** | August (₹7.37 lakh) |
| **Top Occasion by Revenue** | Anniversary (₹6.75 lakh, 19.2%) |
| **Top Category by Revenue** | Colors (₹10.06 lakh, 28.6%) |
| **Top Product by Revenue** | Magnam Set (₹1.22 lakh) |
| **Top Cities by Orders** | Bareilly & Ghaziabad (9 orders each) |

> ⚠️ **Headline insight:** FNP's revenue is highly **seasonal** — four months (Feb, Mar, Aug, Nov) bring **68.2%** of annual revenue, while the other eight months stay flat at roughly ₹1–1.6 lakh each. Festival products sell only in their own season, so growth depends on strengthening year-round occasions such as Anniversary and Birthday.

---

## 📂 Repository Structure

```
ferns-and-petals-sales-analysis/
├── data/
│   ├── customers.csv                 # Raw customer data (100 rows × 7 cols)
│   ├── orders.csv                    # Raw order data (1,000 rows × 10 cols)
│   └── products.csv                  # Raw product data (70 rows × 5 cols)
├── Ferns_Petals.xlsx                 # Final workbook (data + pivots + dashboard)
├── docs/
│   ├── Ferns_Petals_Project_Documentation.docx   # Formal report (Word)
│   └── DOCUMENTATION.md                          # Data dictionary & method (Markdown)
├── images/
│   └── *.png                         # Chart exports used in this README
└── README.md
```

> 💡 Place this README, the `images/` folder, your CSVs and the workbook in the structure above when you push to GitHub so all image links resolve correctly.

---

## 🗃 Datasets

| Dataset | Rows | Columns | Description |
|---|---:|---:|---|
| `customers.csv` | 100 | 7 | Customer ID, name, city, contact number, email, gender, address |
| `orders.csv` | 1,000 | 10 | Order ID, customer ID, product ID, quantity, order & delivery date/time, delivery location, occasion |
| `products.csv` | 70 | 5 | Product ID, name, category, price (INR), occasion |
| **Orders table (enriched in Excel)** | **1,000** | **20** | 10 raw columns + 10 derived columns |

**Relationship model:**

```
Customers ──(Customer_ID)──▶ Orders ◀──(Product_ID)── Products
```

---

## 🧹 Data Preparation Summary

Every table was profiled and validated before analysis.

| Check | Result |
|---|---|
| Missing values | None in any of the three tables |
| Duplicate keys (`Customer_ID`, `Order_ID`, `Product_ID`) | None |
| Orphan foreign keys (orders without a valid customer / product) | None — all 1,000 orders match |
| Occasion consistency (order vs product) | 0 mismatches |
| Customer coverage | All 100 customers placed at least one order (3–17 orders each) |
| Data types | Dates, times, numbers and text IDs set correctly; `Contact_Number` kept as text |

**Derived columns added to the `Orders` table:**

| Field | Formula / Logic |
|---|---|
| `Month Name` | Month of `Order_Date` |
| `Day (Order_Date)` | Weekday of `Order_Date` |
| `Hour [Order_Time]` | Hour (0–23) of order time |
| `Hour [Delivery_Time]` | Hour (0–23) of delivery time |
| `DaysDiff_Order&Delivery` | `Delivery_Date − Order_Date` (days) |
| `Price (INR)` | Unit price looked up from `Products` using `Product_ID` |
| `Revenue` | `Quantity × Price (INR)` |

---

## 🔧 What Was Done in Excel

1. **Loaded** the three CSVs into separate sheets and converted each into an Excel **Table** (`Customers`, `Orders`, `Products`).
2. **Added derived columns** (revenue, delivery days, month, weekday, hour) to the `Orders` table.
3. **Built 11 PivotTables** — KPIs, monthly trend, top products, occasion, category, cities, hourly and weekday patterns.
4. **Created 6 PivotCharts** and assembled them with 4 KPI cards on the **Dashboard** sheet.
5. **Added interactivity** with an **Occasion slicer** and **Order_Date / Delivery_Date timelines**.
6. **Tested the relationship** between quantity and delivery time using `CORREL(Orders[Quantity], Orders[DaysDiff_Order&Delivery])`.

---

## 🖥 Dashboard Overview

| Component | Content |
|---|---|
| **KPI cards** | Total Revenue · Avg. Customer's Spending · Avg. Order-Delivery Time · Total Orders |
| **Charts** | Revenue by Occasions · Revenue by Category · Revenue by Month · Top 10 Cities by Order · Revenue by Hour (Order Time) · Top 5 Products by Revenue |
| **Filters** | Occasion slicer · Order_Date timeline · Delivery_Date timeline |

![All Occasions Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
![Anniversary Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Anniversary.png)
![Birthday Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Birthday.png)
![Diwali Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Diwali.png)
![Holi Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Holi.png)
![Raksha Bandhan Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Raksha%20Bandhan.png)
![Valentine's Day Dashboard](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Valentine's%20Day.png)

> ℹ️ The saved workbook has the **Occasion slicer set to Valentine's Day**, so the dashboard opens showing Valentine's figures (₹3,31,930 revenue, 113 orders). Clear the slicer (funnel icon with ✕) to see the full-year numbers used in this README.

---

## 📈 Analysis & Visualizations

All figures are for the full year 2023 with no filters applied.

### 1. Total Revenue
FNP earned **₹35,20,984** (₹35.21 lakh) from **1,000 orders**, giving an average order value of **₹3,521**.

### 2. Average Order & Delivery Time
Orders take **5.53 days** on average to be delivered, ranging from **1 to 10 days**. About 29% of orders arrive within 3 days, while about 32% take 8 days or more — delivery time is spread widely and is hard for customers to predict.

### 3. Monthly Sales Performance
![Monthly Sales Performance](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
Revenue peaks in **August (₹7.37 lakh)** and **February (₹7.05 lakh)**, followed by **March (₹5.12 lakh)** and **November (₹4.49 lakh)**. The peaks match the festival calendar: all Raksha Bandhan revenue falls in August, all Valentine's Day revenue in February, all Diwali revenue in November, and Holi in February–March.

### 4. Top Products by Revenue
![Top 10 Products by Revenue](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
**Magnam Set** (₹1.22 lakh) leads, followed by **Quia Gift** (₹1.14 lakh), **Dolores Gift** (₹1.07 lakh), **Harum Pack** (₹1.02 lakh) and **Deserunt Box** (₹0.98 lakh). Revenue is not concentrated: the top 5 products make up only **15.4%** of total revenue and the top 10 only 28.8%.

### 5. Customer Spending Analysis
The average customer spends **₹35,210** (median ₹34,747), with spend ranging from ₹8,403 to ₹75,029 (customer **C020**). Each customer placed about 10 orders on average — spend is fairly evenly distributed rather than concentrated in a small group.

### 6. Sales Performance of Top 5 Products
![Top 5 Products — Monthly Revenue](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
Each top product earns most of its revenue in a **single month**: Dolores Gift sells only in August (Raksha Bandhan), Harum Pack peaks in March (Holi), Quia Gift in August, Deserunt Box in April and Magnam Set in February. Product performance follows the occasion calendar, not steady year-round demand.

### 7. Top 10 Cities by Number of Orders
![Top 10 Cities by Orders](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
**Bareilly** and **Ghaziabad** lead with 9 orders each, followed by **Bhilai** (8). After that, 14 cities tie at 7 orders, so ranks 4–10 are shared. Orders are spread over 301 delivery locations and the top three cities account for only 2.6% of orders — no single city dominates.

### 8. Order Quantity vs Delivery Time
![Quantity vs Delivery Time](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)
Higher quantities do **not** slow delivery. The correlation between quantity and delivery days is **0.003** (effectively zero), and average delivery time stays between 5.25 and 5.83 days for every quantity from 1 to 5.

### 9. Revenue Comparison Between Occasions
![Revenue by Occasion](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)

| Occasion | Revenue | Orders | Share | Avg. Order Value |
|---|---:|---:|---:|---:|
| 🟢 Anniversary | ₹6,74,634 | 205 | 19.2% | ₹3,291 |
| 🟢 Raksha Bandhan | ₹6,31,585 | 132 | 17.9% | ₹4,785 |
| All Occasions | ₹5,86,176 | 126 | 16.6% | ₹4,652 |
| Holi | ₹5,74,682 | 180 | 16.3% | ₹3,193 |
| Birthday | ₹4,08,194 | 149 | 11.6% | ₹2,740 |
| Valentine's Day | ₹3,31,930 | 113 | 9.4% | ₹2,937 |
| 🔴 Diwali | ₹3,13,783 | 95 | 8.9% | ₹3,303 |

**Anniversary** earns the most revenue; **Raksha Bandhan** has the highest value per order (₹4,785); **Diwali** is the lowest and the biggest growth opportunity.

### 10. Product Popularity by Occasion
![Best-Selling Product by Occasion](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)

| Occasion | #1 by Revenue | #2 by Revenue | #3 by Revenue |
|---|---|---|---|
| Anniversary | Dignissimos Pack (₹90,036) | Provident Pack (₹84,612) | Delectus Gift (₹75,960) |
| Raksha Bandhan | Dolores Gift (₹1,06,624) | Nostrum Box (₹97,656) | Aperiam Box (₹89,880) |
| All Occasions | Magnam Set (₹1,21,905) | Exercitationem Pack (₹96,701) | Expedita Gift (₹88,944) |
| Holi | Harum Pack (₹1,01,556) | Nam Gift (₹91,385) | Et Set (₹69,601) |
| Birthday | Deserunt Box (₹97,665) | Fuga Set (₹80,640) | Voluptatem Box (₹60,116) |
| Valentine's Day | Eius Gift (₹85,904) | Sed Pack (₹67,140) | Iure Gift (₹51,828) |
| Diwali | Aut Box (₹81,057) | Accusantium Set (₹61,149) | Officiis Pack (₹45,696) |

Each occasion has its own leader. The top product by revenue is not always the top by units sold, because prices differ — use revenue leaders for margin-focused promotions and unit leaders for stock planning.

---

## 💰 Category Performance

![Revenue by Category](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/All%20Occasions.png)

| Category | Revenue | Share |
|---|---:|---:|
| 🟢 Colors | ₹10,05,645 | 28.6% |
| Soft Toys | ₹7,40,831 | 21.0% |
| Sweets | ₹7,33,842 | 20.8% |
| Cake | ₹3,29,862 | 9.4% |
| Raksha Bandhan | ₹2,97,372 | 8.4% |
| Plants | ₹2,12,281 | 6.0% |
| 🔴 Mugs | ₹2,01,151 | 5.7% |

**Key takeaway:** Colors, Soft Toys and Sweets together generate about **70%** of revenue, while Plants and Mugs contribute the least.

---

## ✅ Business Recommendations

1. **Plan around the four peaks** — prepare inventory, staff and delivery capacity before Valentine's Day, Holi, Raksha Bandhan and Diwali.
2. **Grow the off-season** — run Anniversary and Birthday campaigns in the flat months (Apr–Jul, Sep, Oct, Dec).
3. **Prioritise key categories** — focus the catalogue and promotions on Colors, Soft Toys and Sweets.
4. **Make delivery predictable** — the 1–10 day spread is wide; a tighter delivery window matters most for date-sensitive gifts.
5. **Time promotions well** — order revenue peaks in the evening (7–8 PM) with a second peak around 6 AM.
6. **Grow Diwali** — lowest revenue and fewest orders, so targeted offers and bundles could unlock growth.
7. **Reward top customers** — introduce loyalty benefits for the highest-spending customers.

---

## ⚠️ Known Data Notes

- **Grand Total in charts:** the dashboard charts currently include the PivotTable *Grand Total* row (a stray "Grand Total" bar/point). Turn Grand Totals off (*PivotTable Design → Grand Totals → Off*) to remove it.
- **KPI label:** the delivery KPI card says *"Avg. Order-Delivery Time (Hours)"* but the value is in **days**.
- **`Order_Time` column:** in the Excel `Orders` sheet it was converted to a whole number (shows `0`/`1`). The exact time is preserved in `orders.csv`, and the `Hour [Order_Time]` column used by the charts is correct.
- **Duplicate product names:** `Quia Gift` and `Error Gift` each belong to two different Product IDs — analyse by `Product_ID` if you need them separated.
- **Ties in Top 10 cities:** 14 cities tie at 7 orders, so ranks 4–10 are not unique.
- **Synthetic data:** the dataset appears to be synthetic, so conclusions are illustrative.

---

## 🛠 Tech Stack

| Tool | Purpose |
|---|---|
| **Microsoft Excel** | Data preparation, analysis and dashboard |
| **Excel Tables** | Structured, auto-expanding data ranges |
| **PivotTables & PivotCharts** | Summarisation and visualisation |
| **Slicer & Timelines** | Interactive filtering by occasion and date |
| **Excel Formulas** | `CORREL`, lookups and derived columns |

---

## ▶️ How to Use

```bash
# 1. Clone the repository
git clone https://github.com/mandeepkaur-tech/ferns-and-petals-sales-analysis.git
cd ferns-and-petals-sales-analysis
```

1. Open `Ferns_Petals.xlsx` in **Microsoft Excel 2016 or later** (slicers and timelines need desktop Excel; Excel for the web and Google Sheets will not render them fully).
2. Go to the **Dashboard** sheet and use the slicer and timelines to filter.
3. To refresh after changing data: **Data → Refresh All**.

---

## 📄 Documentation

A full formal write-up — Problem Statement, Dataset Overview, Data Dictionary, Methodology, Answers to all 10 Business Questions, Recommendations and Conclusion — is available in [`docs/Ferns_Petals_Project_Documentation.docx`](https://github.com/mandeepkaur-tech/Ferns-Petals-Sales-Analysis-Excel/blob/main/Project_Documentation.docx).

---

## 👤 Author & Contact

**Mandeep Kaur**

- 📧 Email: [kaur.mandeeep.08@gmail.com]
- 💼 LinkedIn: [linkedin.com/in/-mandeep-kaur-](https://www.linkedin.com/in/-mandeep-kaur-/)
- 🐙 GitHub: [github.com/mandeepkaur-tech](https://github.com/mandeepkaur-tech)

Feel free to reach out for questions, feedback, or collaboration on this project.

<div align="center">

**FERNS & PETALS · SALES ANALYSIS DASHBOARD**

</div>
