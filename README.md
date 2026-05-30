💰 FinVest – Personal Finance & Budget Management System
📌 Overview

FinVest is a comprehensive personal finance management platform designed to help individuals track income, expenses, budgets, savings goals, and financial insights from a single dashboard.

The application provides a secure and user-friendly environment where users can manage their financial activities, monitor spending habits, set budget limits, receive alerts, analyze financial trends, and generate reports for better decision-making.

The system follows a modular architecture with role-based access control, transaction auditing, financial analytics, notification management, and administrative monitoring capabilities.

🎯 Project Objective

Managing personal finances often requires maintaining multiple spreadsheets, manual calculations, and scattered records. FinVest addresses these challenges by providing:

Centralized financial management
Real-time income and expense tracking
Budget planning and monitoring
Savings goal management
Financial analytics and reporting
Notification and reminder services
Secure user authentication and authorization
Administrative monitoring and audit capabilities
🏗️ System Workflow
User Registration/Login
          │
          ▼
 Authentication & Authorization
          │
          ▼
     User Dashboard
          │
 ┌────────┼────────┐
 ▼        ▼        ▼
Income  Expense  Budgets
Entry    Entry   Planning
 │         │        │
 └────┬────┴────────┘
      ▼
 Transaction Processing
      ▼
 Account Balance Update
      ▼
 Analytics Engine
      ▼
 Dashboard & Reports
      ▼
 Notifications & Alerts
      ▼
 Savings Goal Tracking
🔄 Application Workflow
Step 1: User Authentication
User registers with email and password.
Passwords are securely hashed before storage.
Session management maintains authenticated access.
Role-based access control differentiates users and administrators.
Step 2: Account Initialization

After successful registration:

Default account is created.
User settings are initialized.
Default categories are generated.
Notification preferences are configured.
Step 3: Transaction Management

Users can:

Add Income
Add Expenses
Categorize Transactions
Edit Transactions
Delete Transactions
Import Transactions from CSV

Each transaction automatically updates account balances.

Step 4: Budget Planning

Users create monthly budgets for specific categories.

System continuously:

Calculates current spending
Compares against budget limits
Generates warning notifications
Detects overspending situations
Step 5: Analytics Processing

Financial data is aggregated to generate:

Monthly Income
Monthly Expenses
Net Savings
Spending Trends
Category Distribution
Budget Utilization
Step 6: Reporting

Reports are generated from historical transaction data and include:

Income Reports
Expense Reports
Savings Reports
Budget Reports
Trend Analysis
Step 7: Notifications

The notification engine generates alerts for:

Budget Thresholds
Budget Exceeded Events
Goal Deadlines
Important Financial Events
Step 8: Savings Goals

Users can:

Create Savings Goals
Track Progress
Monitor Deadlines
Receive Goal Notifications
📋 Core Modules
Module	Description
Authentication & Authorization	User registration, login, session management, role control
Transaction Management	Income and expense recording and maintenance
Category Management	Custom financial categories
Dashboard Module	Financial overview and analytics
Budget Planning	Budget creation and monitoring
Reports & Analytics	Financial reporting and visualization
Reminder & Notifications	Alert generation and management
Data Export & Backup	CSV and JSON exports
Savings & Goals	Goal tracking and monitoring
Account Management	User profile and settings
Admin & System Module	Administrative controls and monitoring
👤 User Stories Implemented
Authentication
User Registration
User Login
Password Reset
Session Management
Financial Management
Add Income
Add Expenses
Categorize Transactions
Edit Transactions
Delete Transactions
Transaction Filtering
Budgeting
Create Monthly Budgets
Budget vs Actual Spending
Threshold Notifications
Analytics
Income vs Expense Comparison
Category-wise Spending Analysis
Average Monthly Spending
Financial KPIs
Savings
Create Savings Goals
Track Goal Progress
Goal Deadline Notifications
Data Management
CSV Transaction Import
CSV Export
JSON Backup
Administration
User Monitoring
Activity Logs
Session Tracking
User Blocking/Unblocking
🛠️ Technology Stack
Backend
Python 3.11+
Flask
SQLAlchemy ORM
Flask-Login
Flask-WTF
Alembic
APScheduler
Frontend
HTML5
CSS3
Bootstrap 5
JavaScript
Jinja2 Templates
Chart.js
Database
MySQL 8
Security
Password Hashing
Session Management
CSRF Protection
Role-Based Access Control
Audit Logging
📂 Project Structure
FinVest/
│
├── app/
│   ├── routes/
│   ├── models/
│   ├── services/
│   ├── templates/
│   ├── static/
│   └── utils/
│
├── migrations/
├── requirements.txt
├── config.py
├── run.py
└── README.md
📊 Key Features
Financial Dashboard
Total Balance Overview
Monthly Income
Monthly Expenses
Net Savings
Recent Transactions
Budget Summary
Analytics & Visualization
Income vs Expense Charts
Category-wise Pie Charts
Monthly Trends
Spending Distribution
Budget Monitoring
Budget Limits
Utilization Tracking
Overspending Detection
Notification Engine
Budget Alerts
Goal Alerts
System Notifications
Data Portability
CSV Import
CSV Export
JSON Backup
🔐 Security Features
Password Hashing using Werkzeug
Secure Session Management
CSRF Protection
Role-Based Access Control
Login Attempt Monitoring
Activity Logging
Audit Trails
🚀 Future Enhancements
Bank Account Integration
AI-Based Expense Prediction
Investment Portfolio Tracking
Mobile Application
Email & SMS Notifications
Multi-Currency Conversion
Advanced Financial Forecasting


👨‍💻 Development Team

FinVest Capstone Project

Developed as part of the Infosys Springboard Internship Program.

