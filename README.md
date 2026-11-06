# 📊 Yash Marketing - Dealer Performance & Insights Engine

An enterprise-grade analytics, dealer 360° profiling, and automated debt & churn risk radar web application built for **Yash Marketing**, powered by **Google Apps Script** and **GitHub Cloud**.

---

## 🚀 Core Features & Capabilities

### 1. 💼 Executive Overview & KPI Dashboard
- **Sales Volume & Average Order Value (AOV):** Real-time aggregation of completed orders.
- **Dealers Health Matrix:** Live count of Active (<30 days), At-Risk (30-75 days), and Hibernating (>75 days) accounts.
- **Outstandings & Credit Health:** Total market dues with instant breakdown of critical aging accounts (>45 days overdue).
- **Interactive Charts (Chart.js):**
  - **Monthly Sales Velocity Curve** (Line/Area chart).
  - **Tier Distribution** (Platinum ₹3L+, Gold ₹1L-₹3L, Silver ₹30k-₹1L, Bronze <₹30k).
  - **Territory Revenue Comparison** (Bar chart across beats & routes).

### 2. ⚡ Intelligent Early-Warning Radar
- **Revenue Drop-off / Churn Alerts:** Detects high-value dealers whose order frequency has halted.
- **Credit Limit Breaches:** Highlights dealers whose pending dues exceed their credit limits.
- **Aging Dues Warning:** Pinpoints dealers with 45+ days unpaid credit.
- **Direct 1-Click Actions:**
  - 📲 **1-Click WhatsApp Reminder:** Opens WhatsApp with an auto-drafted polite statement reminder with dealer name and balance.
  - 📞 **1-Click Call Trigger:** Instant dial from mobile or desktop.

### 3. 🔍 Complete 360° Dealer Profile (Dossier)
- Comprehensive modal displaying:
  - Master details (Phone, Area, Route/Beat, Tier Badge, Credit Limit).
  - Recent Order History (Order IDs, items purchased, invoice amounts, stages).
  - Market Visit Logs (Field staff visit timestamps and discussion notes).
  - Follow-up & Action Tracker (Log calls, meetings, and future follow-up dates directly into Google Sheets).

### 4. 🗺️ Territory & Beat Analytics
- Compare performance across areas (e.g. Ghatkopar, Kurla, Mulund, Thane, etc.).
- Active dealer counts vs outstanding dues per beat.

### 5. 📄 Instant Export & Sync
- **1-Click CSV Export:** Export complete dealer performance summaries for offline accounting or Excel analysis.
- **Zero-Storage Online Architecture:** 100% cloud-synced via GitHub REST API and Google Apps Script GitHub Assistant.

---

## 📋 Required Google Sheet Tabs

The app dynamically links to your active spreadsheet. If any tab doesn't exist, the backend engine automatically initializes it with clean headers:

1. **`Dealer Database`**: `DealerPhone`, `DealerName`, `Area`, `Beat`, `Category`, `CreditLimit`, `Latitude`, `Longitude`, `CreatedDate`
2. **`Order Book`**: `Date & Time`, `OrderID`, `ShopName`, `DealerPhone`, `OrderItems`, `TotalAmount`, `Payment Mode`, `Stage`, `DeliveryStatus`
3. **`Outstanding`**: `ShopName`, `WithBillAmount`, `WithoutBillAmount`, `OverdueDays`, `LastPaymentDate`, `CreditStatus`
4. **`Market Visits`**: `Timestamp`, `StaffName`, `ShopName`, `Status`, `RadialDistance`, `Notes`, `PhotoURL`
5. **`Dealer Action Logs`**: `Timestamp`, `ShopName`, `ActionType`, `Notes`, `LoggedBy`, `FollowUpDate`

---

## 🔗 Setup with Google Apps Script GitHub Assistant

1. Open your **Google Apps Script** project (linked to your Yash Marketing Master Sheet).
2. Open the **Google Apps Script GitHub Assistant** browser extension.
3. Configure the extension:
   - **Repository:** `yashmbarlota007/Dealer-Performance-Insights`
   - **Branch:** `main`
4. Click **Pull from GitHub** to instantly sync `Code.gs`, `Index.html`, and `appsscript.json`.
5. Click **Deploy** -> **New Deployment** -> **Web App**:
   - **Execute as:** `Me` (Your Google Account)
   - **Who has access:** `Anyone` (or within your workspace)
6. Access your live Dealer Performance & Insights Engine dashboard!

---

## 🔒 Zero-Storage Online Workflow
This repository is managed 100% online in the cloud using the GitHub REST API. Local disk storage footprint remains 0 KB.
