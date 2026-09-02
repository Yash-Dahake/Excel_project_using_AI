# 🤖 AI-Assisted Sales Performance Dashboard – Excel & VBA

An **AI-assisted, interactive Sales Performance Dashboard** developed using **Microsoft Excel and VBA**. The project transforms raw sales transaction data into meaningful business insights through automated Pivot Tables, KPI cards, charts, Year/Month slicers, and dashboard refresh functionality.

> **🤖 AI-Assisted Project:** AI tools were used during the development process for VBA code generation, debugging, dashboard design guidance, and documentation. The final implementation, testing, customization, and project decisions were performed by the author.

---

## 📌 Project Overview

This project demonstrates how **Excel, VBA, and AI-assisted development** can be combined to create an automated sales analytics dashboard.

The dashboard converts raw transactional sales data into an interactive business intelligence solution that helps analyze:

* Sales performance
* Profitability
* Customer segments
* Product categories
* Regional sales
* Monthly trends
* Discount impact

### Key Features

* 🤖 AI-assisted VBA development
* 📊 Automated dashboard creation
* 🎯 Six KPI cards
* 📈 Six sales visualizations
* 📅 Year and Month slicers
* 🔄 Automated Refresh Dashboard button
* 📋 Pivot Tables as analytical backend
* 🎨 Professional purple-themed dashboard
* ⚡ Automated formatting and layout

---

## 📊 Dashboard KPIs

| KPI                    |      Value |
| ---------------------- | ---------: |
| 💰 Total Sales         | ₹89,72,527 |
| 💵 Total Profit        | ₹12,99,797 |
| 🧾 Total Orders        |      8,683 |
| 📦 Total Quantity Sold |     48,043 |
| 🏷️ Average Discount   |      13.7% |
| 📈 Profit Margin       |      14.5% |

---

## 📈 Dashboard Visualizations

### 1. Sales by Segment

**Chart Type:** Doughnut Chart

Analyzes sales contribution from:

* Consumer
* Corporate
* Home Office

### 2. Sales & Profit Trend

**Chart Type:** Line Chart

Shows the trend of sales and profit over time.

### 3. Monthly Sales Performance

**Chart Type:** Clustered Column Chart

Compares monthly sales performance across different periods.

### 4. Profit vs Discount

**Chart Type:** Column Chart

Analyzes profit and average discount across product sub-categories.

### 5. Sales by Category

**Chart Type:** Horizontal Bar Chart

Compares sales across:

* Technology
* Furniture
* Office Supplies

### 6. Sales by Region

**Chart Type:** Clustered Column Chart

Compares sales across:

* East
* South
* West
* Central

---

## 🎛️ Interactive Slicers

The dashboard includes:

* 📅 **Year Slicer**
* 📆 **Month Slicer**

The slicers are connected to the dashboard Pivot Tables, allowing users to interactively filter the analysis.

Selecting a particular year or month updates the relevant Pivot Table outputs and dashboard visualizations.

---

## 🔄 Dashboard Refresh

A **Refresh Dashboard** button is included in the dashboard.

The VBA automation refreshes the workbook data and Pivot-related components, making it easier to update the dashboard when the underlying dataset changes.

---

## 🤖 AI-Assisted Development

AI was used as a **development assistant** throughout the project.

### AI was used for:

* Generating and improving VBA code
* Debugging VBA errors
* Developing Pivot Table automation
* Creating chart automation logic
* Implementing Year and Month slicers
* Connecting slicers with Pivot Tables
* Improving dashboard layout and formatting
* Generating documentation and README content
* Explaining Excel and VBA concepts during development

### Human Contribution

The project was not simply generated and submitted without modification.

The author was responsible for:

* Preparing and organizing the dataset
* Defining dashboard requirements
* Selecting KPIs and visualizations
* Reviewing and modifying AI-generated VBA code
* Testing the dashboard
* Identifying and fixing implementation issues
* Customizing the dashboard design
* Validating the final results
* Interpreting the business insights

> **AI was used as a development and learning assistant, while the final project implementation and decisions were reviewed and customized by the author.**

---

## 🛠️ Data Structure

The `Data` worksheet contains the following fields:

| Column       | Description             |
| ------------ | ----------------------- |
| Order ID     | Unique order identifier |
| Order Date   | Date of the order       |
| Year         | Order year              |
| Month        | Order month             |
| Day          | Day of the order        |
| Segment      | Customer segment        |
| Category     | Product category        |
| Sub-Category | Product sub-category    |
| Region       | Sales region            |
| Sales        | Sales revenue           |
| Profit       | Profit generated        |
| Quantity     | Quantity sold           |
| Discount     | Discount percentage     |

---

## 📁 Project Structure

| Sheet         | Purpose                                              |
| ------------- | ---------------------------------------------------- |
| **Data**      | Raw sales transaction data                           |
| **Pivot**     | Backend Pivot Tables used for analysis               |
| **Dashboard** | Interactive dashboard with KPIs, charts, and slicers |

### Data Flow

```text
Raw Sales Data
       ↓
      Data
       ↓
   Pivot Tables
       ↓
 Year / Month Slicers
       ↓
 KPIs + Charts
       ↓
Sales Performance Dashboard
```

---

## 📈 Key Business Insights

### 1. Segment Performance

The **Consumer segment** is the largest contributor to overall sales, accounting for more than **52%** of total sales.

Corporate contributes approximately **30%**, while Home Office contributes around **18%**.

### 2. Category Performance

**Technology** is the leading category with approximately **₹45.84L** in sales.

Furniture follows with approximately **₹37.13L**, while Office Supplies contributes approximately **₹6.76L**.

### 3. Regional Performance

| Region  |   Sales |
| ------- | ------: |
| East    | ₹23.21L |
| South   | ₹22.53L |
| West    | ₹22.25L |
| Central | ₹21.73L |

The **East region** records the highest sales among the four regions.

### 4. Profitability

The overall **profit margin is 14.5%**, showing a positive relationship between sales revenue and profitability.

### 5. Discount Analysis

The **Profit vs Discount** analysis helps identify sub-categories where discount levels may influence profitability.

This can support better discount and pricing decisions.

---

## 🎯 Conclusion

The **AI-Assisted Sales Performance Dashboard** demonstrates how traditional Excel analytics can be enhanced through **VBA automation and AI-assisted development**.

The combination of Pivot Tables, KPIs, charts, slicers, and VBA automation provides an efficient way to analyze sales performance and generate actionable business insights.

The project also demonstrates practical experience in using **AI as a development assistant** while maintaining human involvement in requirements, testing, customization, validation, and business interpretation.

---

## 💻 Technology Stack

* **Microsoft Excel**
* **VBA (Visual Basic for Applications)**
* **Pivot Tables**
* **Pivot Charts**
* **Slicers**
* **Excel Formulas**
* **AI-Assisted Development**

---

## ⚙️ How to Run

### 1. Open the Workbook

Open the Excel workbook containing the sales dataset.

### 2. Enable Macros

Make sure Excel macros are enabled.

### 3. Open VBA Editor

Press:

```text
Alt + F11
```

### 4. Insert the VBA Module

Go to:

```text
Insert → Module
```

Paste the complete VBA dashboard code into the module.

### 5. Run the Main Macro

Run:

```text
BuildCompleteSalesDashboard
```

The macro will create the required:

* Pivot Tables
* Dashboard
* KPI cards
* Charts
* Year slicer
* Month slicer
* Refresh button

### 6. Interact With the Dashboard

Use the **Year** and **Month** slicers to filter the dashboard and click **Refresh Dashboard** when the underlying data is updated.

---

## 📸 Dashboard Preview

Add your dashboard screenshot here:

```markdown
![AI-Assisted Sales Performance Dashboard](images/dashboard.png)
```

---

## 👤 Author

### Yash Rajesh Dahake

**Aspiring Data Analyst**

**Skills:** Excel | VBA | Power BI | Python | SQL | Data Analysis

---

⭐ **AI-assisted project developed for learning, portfolio building, and practical data analytics experience.**
