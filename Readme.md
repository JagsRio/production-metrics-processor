# Production Metrics Reporter Windows Service

**VB.NET • Windows Service • Automated Reporting • System Monitoring**

A sanitized, portfolio‑ready Windows Service that demonstrates how to automate daily operational reporting, monitor system activity, and aggregate performance metrics in an enterprise environment.  
This version removes all proprietary identifiers, schemas, and infrastructure details while preserving the architectural patterns and engineering quality of the original production service.

---

## 🚀 Overview

The **Production Metrics Reporter** is a Windows NT Service that performs scheduled data monitoring and generates automated reports summarizing:

- Daily processing volume  
- Stale client activity  
- Performance exceptions (e.g., low success rates)  

It is designed to run unattended, execute during a configurable daily window, and log all activity for diagnostics.

---

## 🔧 Key Features

### **Windows Service Architecture**
- Implements `ServiceBase`
- Clean `OnStart` / `OnStop` lifecycle
- Uses `System.Timers.Timer` for scheduled execution

### **Automated Daily Reporting**
- Executes once per day at a configurable hour
- Generates three types of reports:
  - **Volume Summary**
  - **Stale Activity Report**
  - **Performance Exceptions Report**

### **Database Querying & Aggregation**
- Uses ADO.NET (`SqlConnection`, `SqlDataAdapter`)
- Retrieves client activity and performance metrics
- Aggregates and formats data into tabular reports

### **Email Alerting (Sanitized)**
- Production version uses SMTP
- In this sanitized version, email alerts are logged to a file instead of being sent, preserving functionality without exposing sensitive infrastructure.
- Portfolio version logs email events instead of sending

### **Configuration‑Driven Behavior**
- Monitoring interval
- Database connection string
- Logging path
- Daily trigger hour

### **Robust Logging & Error Handling**
- Logs all service events and report generation steps
- Captures and logs exceptions without interrupting service execution

---

## 🏁 Getting Started
1. Compile the solution in Visual Studio.
2. Install the Windows Service via `sc create` or `InstallUtil.exe`.
3. Configure `app.config` for logging path, database connection, and daily trigger hour.
4. Start the service via Services.msc or `net start`.

--- 

## 🧠 Technical Highlights
- Demonstrates enterprise‑grade background processing
- Clean separation of responsibilities:
  - Data retrieval
  - Report generation
  - Email/report dispatch
  - Logging
- Uses defensive programming patterns
- Sanitized SQL queries for safe public sharing
- Shows how to build tabular text reports using `StringBuilder`