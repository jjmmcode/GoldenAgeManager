# Golden Age Manager

**Discord Community Management & Automation Bot**

## Overview

Golden Age Manager is a Python-based Discord automation bot developed to streamline administrative workflows for a gaming community.

The project automates application processing, synchronizes user roles, manages character and mentor registries, and generates administrative reports.

## Core Features

### Application Management
- Automated processing of character, mentor, and transformation applications.
- Management of approved and rejected applications.
- Duplicate processing prevention.

### Role & Access Management
- Automatic synchronization of Discord roles.
- Role-based access control.
- Restoration of eligible user permissions.

### Character & Mentor Management
- Character registration and race tracking.
- Automated mentor directories.
- Management of available mentor abilities.

### Reporting & Automation
- Excel report generation using openpyxl.
- Word document generation using python-docx.
- Automated community maintenance.
- Persistent application state using JSON.

## Tech Stack

| Technology | Purpose |
|---|---|
| Python | Core application logic |
| discord.py | Discord API integration |
| asyncio | Asynchronous operations |
| JSON | Persistent state management |
| openpyxl | Excel report generation |
| python-docx | Word document generation |
| python-dotenv | Environment configuration |

## Architecture

The application uses an event-driven architecture built around Discord events and asynchronous operations.

Its main responsibilities include:

1. Receiving and processing Discord events.
2. Validating applications and administrative actions.
3. Managing community roles and permissions.
4. Maintaining character and mentor information.
5. Generating reports and maintaining persistent processing records.

## Project Structure

```text
GoldenAgeManager/
├── bot.py
├── requirements.txt
├── README.md
├── .gitignore
└── .gitattributes
```

## Security

The Discord bot token is loaded from environment variables rather than being hardcoded into the application.

Sensitive credentials and configuration files should never be committed to the repository.

## Author

**Joel Medina**

Junior Data Engineer & Software Developer

[GitHub](https://github.com/jjmmcode) | [Portfolio](https://jjmmcode-portafolio.vercel.app/)
