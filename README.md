# 💰 FinVest – Personal Finance Management System

A comprehensive web-based Personal Finance Management Platform built using Flask and MySQL that helps users track income, expenses, budgets, savings goals, and financial insights through an interactive dashboard.

---

## 📖 Overview

FinVest is designed to simplify personal financial management by providing a centralized platform for managing day-to-day financial activities.

The application allows users to:

- Track income and expenses
- Create and manage budgets
- Set financial goals
- Analyze spending habits
- Generate reports
- Receive financial alerts and reminders
- Monitor overall financial health

The system follows a modular architecture with secure authentication, role-based access control, financial analytics, and administrative monitoring capabilities.

---

## 🎯 Project Objective

Managing personal finances often involves maintaining spreadsheets and manual records.

FinVest solves this problem by providing:

- Centralized Financial Management
- Real-Time Income & Expense Tracking
- Budget Planning & Monitoring
- Savings Goal Management
- Financial Analytics & Reporting
- Notifications & Alerts
- Secure User Authentication
- Administrative Monitoring

---

# 🔄 Application Workflow

```text
User Registration/Login
          │
          ▼
 Authentication & Authorization
          │
          ▼
      Dashboard
          │
 ┌────────┼────────┐
 ▼        ▼        ▼
Income  Expenses  Budgets
Entry    Entry    Planning
 │         │         │
 └────┬────┴─────────┘
      ▼
 Transaction Processing
      ▼
 Account Balance Update
      ▼
 Analytics Engine
      ▼
 Reports & Dashboard
      ▼
 Notifications & Alerts
      ▼
 Goal Tracking
```

---

# 📋 Core Modules

| Module | Description |
|----------|-------------|
| Authentication & Authorization | User registration, login, password reset, session management |
| Transaction Management | Record and manage income and expense transactions |
| Category Management | Create and manage custom transaction categories |
| Dashboard Module | Financial summary, charts, and KPIs |
| Budget Planning | Monthly budget creation and tracking |
| Reports & Analytics | Spending analysis and financial reports |
| Reminder & Notifications | Budget alerts and reminders |
| Data Export & Backup | CSV import/export and backups |
| Savings & Goals | Financial goal creation and tracking |
| Account Management | User profile and settings management |
| Admin & System Module | User monitoring and system administration |

---

# 👤 User Roles

## End User

Users can:

- Register and Login
- Manage Income
- Manage Expenses
- Create Budgets
- Set Savings Goals
- View Reports
- Receive Notifications
- Export Financial Data

---

## Administrator

Administrators can:

- Manage Users
- Monitor System Activity
- View Audit Logs
- Manage Categories
- Review Usage Metrics

---

# ✨ Features

## 🔐 Authentication

- User Registration
- Secure Login
- Password Reset
- Session Management
- Role-Based Access Control

---

## 💵 Transaction Management

- Add Income
- Add Expenses
- Edit Transactions
- Delete Transactions
- Transaction History
- Date Range Filtering

---

## 📂 Category Management

- Income Categories
- Expense Categories
- Custom Categories

---

## 📊 Dashboard & Analytics

- Total Balance
- Monthly Income
- Monthly Expenses
- Net Savings
- Recent Transactions
- Financial KPIs

### Analytics

- Income vs Expense Comparison
- Category-wise Spending Analysis
- Monthly Spending Trends
- Average Monthly Spending

---

## 💰 Budget Planning

- Monthly Budget Creation
- Budget Monitoring
- Budget Utilization Tracking
- Overspending Detection

### Budget Alerts

- 80% Threshold Warning
- 100% Budget Exceeded Alert

---

## 🎯 Savings Goals

Users can:

- Create Goals
- Track Progress
- Set Target Amount
- Set Deadlines
- Monitor Goal Completion

---

## 🔔 Notification System

Notifications are generated for:

- Budget Threshold Reached
- Budget Exceeded
- Goal Deadlines
- Financial Reminders

---

## 📁 Data Management

- CSV Import
- CSV Export
- Backup Support

---

# 🏗️ Database Design

The application is built using the following database models:

### User

Stores user account information.

### Account

Stores user financial accounts and balances.

### Category

Stores income and expense categories.

### Transaction

Stores financial transaction records.

### Budget

Stores category-wise monthly budgets.

### Goal

Stores savings goals and progress.

### UserSettings

Stores user preferences and configurations.

---

# 🛠️ Technology Stack

## Backend

- Python 3.11+
- Flask
- SQLAlchemy ORM
- Flask-Login
- Flask-WTF
- Alembic
- APScheduler

---

## Frontend

- HTML5
- CSS3
- Bootstrap 5
- JavaScript
- Jinja2 Templates
- Chart.js

---

## Database

- MySQL 8

---

## Development Tools

- Git
- GitHub
- VS Code
- Linux

---

# 🔐 Security Features

- Password Hashing
- Session Management
- CSRF Protection
- Role-Based Access Control
- Audit Logging
- Secure Authentication Flow

---

# 📂 Project Structure

```bash
FinVest/
│
├── backend/
│
├── app/
│   ├── models/
│   ├── routes/
│   ├── services/
│   ├── templates/
│   ├── static/
│   └── utils/
│
├── migrations/
│
├── requirements.txt
├── config.py
├── run.py
└── README.md
```

---

# 🚀 Installation

### Clone Repository

```bash
git clone https://github.com/your-username/FinVest.git
cd FinVest
```

### Create Virtual Environment

```bash
python -m venv venv
```

### Activate Environment

Windows:

```bash
venv\Scripts\activate
```

Linux/Mac:

```bash
source venv/bin/activate
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Configure Database

Update database settings inside:

```python
config.py
```

### Run Migrations

```bash
flask db upgrade
```

### Start Application

```bash
python run.py
```

---


# 👨‍💻 Developer

**Sudhanshu More**

- GitHub: https://github.com/sudhanshumore10
- LinkedIn: Add Your LinkedIn Profile

---

# 📄 License

This project was developed as part of the Infosys Springboard Internship Program and is intended for educational and learning purposes.
