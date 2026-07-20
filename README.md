# EngFlow
Enterprise workflow automation platform designed to optimize engineering operations by automating request management, tracking process lifecycles, integrating data workflows, and reducing manual effort through intelligent system orchestration.
---

## Overview

Engineering Workflow Automation is a desktop application developed in Python to streamline engineering and operational processes through automation.

The platform provides specialized modules for request management, workflow execution, spreadsheet integration, and process tracking, allowing teams to reduce repetitive manual activities and improve operational efficiency.

Designed with scalability, maintainability, and user experience in mind, the application serves as a centralized environment for managing engineering workflows and automation tasks.

---

## Application Preview

### Secure Authentication

The platform provides a secure authentication layer that manages access to automation modules and workflow management features.



---

### Phoenix Module

The Phoenix module is responsible for engineering request creation, lifecycle management, and process tracking.

Features include:

- Request registration
- Workflow navigation
- Part Number management
- Request lifecycle monitoring
- Operational task automation



---

### Pegasus Module

The Pegasus module focuses on spreadsheet integration, workflow execution, and operational data management.

Features include:

- Spreadsheet synchronization
- Automated data updates
- Workflow execution
- Process maintenance
- Operational support activities



---

## Key Features

### Workflow Automation

- Automated engineering processes
- Workflow execution and monitoring
- Reduced manual effort
- Standardized operations

### Request Management

- Request creation
- Request tracking
- Lifecycle management
- Status monitoring

### Spreadsheet Integration

- Excel integration
- Automated updates
- Data synchronization
- Process control

### User Experience

- Modern dark-themed interface
- Modular navigation
- Centralized workflow access
- Persistent configuration storage

### Reliability

- Error handling
- Validation mechanisms
- Historical records
- Secure authentication

---

## Technology Stack

### Programming Language

- Python

### Automation

- Playwright

### Desktop Interface

- Tkinter

### Data Processing

- OpenPyXL
- JSON

### Productivity Tools

- Excel
- Spreadsheet-based workflows

---

## Architecture

```text
┌─────────────────────────┐
│          User           │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│   Authentication Layer  │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│   Desktop Application   │
└────────────┬────────────┘
             │
   ┌─────────┴─────────┐
   │                   │
   ▼                   ▼
┌───────────┐   ┌───────────┐
│  Phoenix  │   │  Pegasus  │
└─────┬─────┘   └─────┬─────┘
      │               │
      └───────┬───────┘
              │
              ▼
┌─────────────────────────┐
│   Workflow Controller   │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│ Spreadsheet Integration │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│ Historical Data Storage │
└─────────────────────────┘
```

---

## Project Objectives

This project was developed with the following objectives:

- Automate repetitive engineering tasks
- Reduce manual workload
- Improve operational visibility
- Increase productivity
- Standardize workflow execution
- Improve request tracking
- Centralize process management
- Create a scalable automation foundation

---

## Software Engineering Principles

The application follows several software engineering best practices:

- Modular architecture
- Separation of concerns
- Maintainable codebase
- Reusable components
- Structured error handling
- Configuration management
- Version control
- Documentation-first approach

---

## Installation

### Clone the Repository

```bash
git clone https://github.com/yourusername/engineering-workflow-automation.git
```

### Navigate to the Project Folder

```bash
cd engineering-workflow-automation
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Run the Application

```bash
python main.py
```

---

## Project Structure

```text
engineering-workflow-automation
│
├── assets
│   └── screenshots
│       ├── login.png
│       ├── phoenix-module.png
│       └── pegasus-module.png
│
├── config
│
├── src
│   ├── services
│   ├── ui
│   ├── automation
│   ├── storage
│   └── utils
│
├── requirements.txt
├── README.md
├── .gitignore
└── LICENSE
```

---

## Future Improvements

Planned improvements include:

- REST API integration
- Database support
- Cloud synchronization
- Real-time notifications
- Analytics dashboard
- User management
- Role-based permissions
- Web version of the platform

---

## Author

### Gabriell Manssur Girotto

Software Engineering Student with a focus on:

- Process Automation
- Software Development
- Workflow Optimization
- Enterprise Solutions
- Python Development

---

## License

This project is licensed under the MIT License.

---

*Engineering Workflow Automation was developed to demonstrate practical application of software engineering principles through workflow automation, operational efficiency, and process management.*
