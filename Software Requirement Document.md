# Software Requirements Document (SRD)

# Improvement Project Monitoring Dashboard
Version: 1.0
Prepared By: Muhammad Nur Abdi
Platform: Streamlit
Database Source: MASTER_PROGRESS.xlsx

---

# 1. Purpose

The purpose of this application is to provide centralized monitoring and performance tracking of Improvement Projects across all departments within Site CAM.

The application consolidates project progress data from multiple departmental improvement plans and presents real-time project performance, achievement, delays, and management insights through an interactive dashboard.

The dashboard will be used by:

- General Manager
- Department Head
- Section Head
- Project PIC
- Continuous Improvement Team

---

# 2. Business Background

Currently, each department maintains its own Improvement Project Plan using separate Excel files.

Challenges:

- Difficult to monitor overall project status.
- Difficult to identify delayed projects.
- Difficult to compare department performance.
- No centralized visibility for management.
- Manual reporting consumes significant time.

The dashboard will automate consolidation and provide a single source of truth for all improvement projects.

---

# 3. Objectives

The system shall:

- Consolidate all project improvement files into one master database.
- Monitor project performance daily.
- Compare Plan vs Actual progress.
- Measure achievement percentage.
- Identify delayed actions.
- Evaluate PIC performance.
- Provide management-level reporting.
- Support decision making.

---

# 4. Scope

## Included

- Executive Dashboard
- Department Monitoring
- Project Monitoring
- Action Plan Monitoring
- PIC Performance
- Factor Analysis
- Category Analysis
- Delay Analysis

## Excluded

- Project creation
- Project editing
- User approval workflow
- Notification system
- Mobile application

---

# 5. Data Sources

Primary Source:

MASTER_PROGRESS.xlsx

Sheets:

- PROJECT_MASTER
- PROGRESS_MASTER
- PROJECT_DAILY_PROGRESS
- SUMMARY_PROJECT
- LATEST_STATUS
- EXECUTIVE_SUMMARY
- DEPARTMENT_PERFORMANCE
- PIC_PERFORMANCE
- FACTOR_PERFORMANCE
- CATEGORY_PERFORMANCE
- DELAY_ANALYSIS
- PROJECT_HEALTH
- PROJECT_RANKING

---

# 6. User Roles

## Executive

Access:

- Executive Dashboard
- Department Performance
- Project Performance

Permissions:

- Read Only

---

## Department Head

Access:

- Department Dashboard
- Project Monitoring
- Action Plan Monitoring

Permissions:

- Read Only

---

## Improvement Team

Access:

- Full Dashboard

Permissions:

- Read Only

---

# 7. Functional Requirements

---

## FR-001 Executive Dashboard

Display executive summary KPIs.

KPIs:

- Total Departments
- Total Projects
- Total Action Plans
- Average Achievement
- Healthy Projects
- Watchlist Projects
- Critical Projects

Data Source:

EXECUTIVE_SUMMARY

---

## FR-002 Department Performance Dashboard

Display department performance ranking.

Metrics:

- Total Projects
- Total Actions
- Average Achievement
- Project Health

Visualization:

- Bar Chart
- Ranking Table

Data Source:

DEPARTMENT_PERFORMANCE

---

## FR-003 Project Monitoring

Display detailed project performance.

Filters:

- Department
- Project

Metrics:

- Plan Progress
- Actual Progress
- Achievement
- Active Activities
- Total Activities

Visualization:

- Plan vs Actual Trend Chart

Data Source:

PROJECT_DAILY_PROGRESS

---

## FR-004 Project Health Monitoring

Display overall project health status.

Categories:

- Healthy
- Watchlist
- Critical

Rules:

Healthy:
Achievement >= 100%

Watchlist:
90% <= Achievement < 100%

Critical:
Achievement < 90%

Data Source:

PROJECT_HEALTH

---

## FR-005 Action Plan Monitoring

Display all action plans.

Columns:

- Department
- Project
- Factor
- Category
- Action Plan
- PIC
- Plan Start
- Due Date
- Achievement
- Status

Filters:

- Department
- Project
- PIC
- Factor
- Category

Data Source:

LATEST_STATUS

---

## FR-006 PIC Performance

Display PIC performance ranking.

Metrics:

- Total Assigned Actions
- Average Achievement

Visualization:

- Ranking Table
- Bar Chart

Data Source:

PIC_PERFORMANCE

---

## FR-007 Factor Analysis

Display performance by root cause category.

Factors:

- Man
- Method
- Machine
- Material
- Measurement
- Environment

Metrics:

- Average Achievement

Visualization:

- Horizontal Bar Chart

Data Source:

FACTOR_PERFORMANCE

---

## FR-008 Category Analysis

Display performance by category.

Categories:

- Quick Action
- Long Term
- Development

Metrics:

- Average Achievement

Visualization:

- Bar Chart

Data Source:

CATEGORY_PERFORMANCE

---

## FR-009 Delay Analysis

Display delayed action plans.

Metrics:

- Delay Days
- PIC
- Department
- Project

Visualization:

- Ranking Table

Color Rules:

Delay > 10 Days = Red

Delay > 5 Days = Orange

Delay <= 5 Days = Yellow

Data Source:

DELAY_ANALYSIS

---

## FR-010 Project Ranking

Display best and worst performing projects.

Metrics:

- Achievement %

Visualization:

- Ranking Table
- Top 10 Projects
- Bottom 10 Projects

Data Source:

PROJECT_RANKING

---

# 8. Dashboard Structure

## Page 1

Executive Dashboard

Components:

- KPI Cards
- Project Health Pie Chart
- Department Ranking
- Project Ranking

---

## Page 2

Project Monitoring

Components:

- Project Filters
- Plan vs Actual Trend
- Project KPI Summary
- Latest Status Table

---

## Page 3

Action Plan Monitoring

Components:

- Detail Table
- Advanced Filters

---

## Page 4

PIC Performance

Components:

- PIC Ranking
- Achievement Chart

---

## Page 5

Factor Analysis

Components:

- Fishbone Performance Chart

---

## Page 6

Category Analysis

Components:

- Category Comparison Chart

---

## Page 7

Delay Analysis

Components:

- Delay Ranking Table
- Delay Summary

---

# 9. Non-Functional Requirements

## Performance

Dashboard loading time:

< 5 seconds

Target Dataset:

- 20+ Projects
- 200+ Action Plans
- 20,000+ Daily Records

---

## Availability

Application shall be available during working hours.

Target Availability:

99%

---

## Usability

Requirements:

- Simple navigation
- Responsive layout
- Interactive filters
- Export capability

---

# 10. Technical Requirements

Frontend:

- Streamlit

Backend:

- Python

Libraries:

- pandas
- plotly
- openpyxl
- numpy

File Format:

- Excel (.xlsx)

Deployment:

- Streamlit Community Cloud
or
- Internal Server

---

# 11. Future Enhancements

Phase 2:

- Automated file upload
- User authentication
- Email notification
- WhatsApp notification
- Project forecasting
- AI-generated project insights
- PowerPoint report generation

---

# 12. Success Criteria

The project shall be considered successful when:

- 100% departmental projects are consolidated.
- Management can monitor all projects from a single dashboard.
- Manual reporting effort is reduced by at least 80%.
- Project status can be identified within 1 minute.
- Delayed projects can be proactively managed.

---
END OF DOCUMENT