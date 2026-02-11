# Modern SharePoint Online for Administrators (2026 Aligned)

A complete **3-day instructor-led training course** for SharePoint Online administrators, fully aligned with current Microsoft 365, Entra ID, Microsoft Purview, Microsoft Search, Power Platform, and Graph-based administration. This course is a **modern replacement for legacy M55238B**, following the Microsoft Official Curriculum (MOC) style.

---

## Course Details

| | |
|---|---|
| **Duration** | 3 Days (+ 1 optional module) |
| **Level** | Intermediate |
| **Audience** | SharePoint Online Administrators, Microsoft 365 Administrators, IT Professionals |
| **Prerequisites** | Basic Microsoft 365 administration and PowerShell knowledge |
| **Lab Scenario** | Project Northwind Intranet Modernization (shared tenant, 10 participants + trainer) |
| **Certification Alignment** | MS-102: Microsoft 365 Administrator · SC-300: Identity and Access Administrator |

---

## Repository Structure

### 📘 Course Syllabus

| File | Description |
|---|---|
| `modern_share_point_online_administrator_2026_aligned_syllabus.md` | Full course syllabus with module outlines, topics, labs, and completion outcomes |

### 📂 Module Folders (`module 1/` – `module 12/`)

Each module folder contains the teaching and lab materials for that module:

| File | Purpose |
|---|---|
| `Module-XX-Slides.md` | Slide outline and trainer talk-track for the module |
| `Module-XX-Slides.pptx` | Generated PowerPoint presentation (widescreen, branded design) |
| `Module-XX-Student-Guide.md` | Student-facing reading material with concepts, procedures, and references |
| `Lab-XX-*.md` | Hands-on lab instructions with step-by-step exercises and deliverables |
| `README.md` | Module-level overview and file listing |

### 📅 Day-by-Day Module Map

#### Day 1 — Tenant Foundations & Site Management

| Module | Title |
|---|---|
| **Module 1** | Introduction to Microsoft 365 and SharePoint Online |
| **Module 2** | Identity, Access, and External Sharing |
| **Module 3** | Working with Site Collections |

#### Day 2 — Information Architecture, Search & Customisation

| Module | Title |
|---|---|
| **Module 4** | Permissions and Collaboration Model |
| **Module 5** | Managing Metadata and the Term Store |
| **Module 6** | Search in SharePoint Online and Microsoft Search |
| **Module 7** | Apps and Customisation in SharePoint Online |

#### Day 3 — Governance, Compliance & Automation

| Module | Title |
|---|---|
| **Module 8** | Content Governance and Compliance with Microsoft Purview |
| **Module 9** | OneDrive for Business Administration |
| **Module 10** | Administration and Automation with PowerShell |
| **Module 11** | Operations at Scale: Monitoring, Lifecycle & External Sharing Governance |
| **Module 12** | *(Optional)* Workflow Automation with Power Automate and Power Apps |

### 🎭 Scenario (`scenario/`)

| File | Description |
|---|---|
| `Lab-Scenario-Overview.md` | Shared-tenant lab scenario — "Project Northwind Intranet Modernization". Defines participant naming conventions (`NW-Pxx-*`), operating rules, and environment baseline used across all labs. |

### 🎒 Participant Packs (`participant-packs/`)

Pre-generated materials for each participant (P01–P10) and the trainer:

| Folder / File | Description |
|---|---|
| `P01/` – `P10/` | Per-participant packs containing a 1-page handout (DOCX/PDF), a 3-slide quick brief (PPTX), a tracker workbook (XLSX), and copy/paste text templates for labs |
| `TRAINER/` | Trainer-only pack with runbooks and answer keys |
| `MANIFEST.md` | Inventory of all generated participant pack contents |

### ⚙️ Scripts (`scripts/`)

| File | Description |
|---|---|
| `Provision-TrainingUsers-P01-P10.ps1` | PowerShell script to provision the 10 training user accounts in the shared tenant |
| `Reassign-Licenses-TrainingUsers.ps1` | PowerShell script to reassign Microsoft 365 licences to training users |

---

## Course Completion Outcomes

After completing this course, students will be able to:

- Administer SharePoint Online using modern tools and admin centers
- Secure collaboration using Microsoft Entra ID and Purview
- Design scalable site, permission, and metadata architectures
- Manage search and content discovery with Microsoft Search
- Implement compliance controls (retention, sensitivity labels, DLP)
- Automate administrative tasks with SharePoint Online Management Shell and Microsoft Graph PowerShell
- Monitor incidents, manage site lifecycle, and govern external sharing at scale
- *(Optional)* Build workflow automation with Power Automate and Power Apps
