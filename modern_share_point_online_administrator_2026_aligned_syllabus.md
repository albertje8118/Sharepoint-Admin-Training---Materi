# Course: Modern SharePoint Online for Administrators (3-Day, 2026 Aligned)

This course is a **modern replacement for legacy M55238B**, fully aligned with **current SharePoint Online, Microsoft 365, Entra ID, Microsoft Purview, Microsoft Search, Power Platform, and Graph-based administration**.

The structure follows the **Microsoft Official Curriculum (MOC)** style: **Modules + Topics + Labs**, optimized for **3 full training days**.

---

## Course Overview

**Duration:** 3 Days  
**Level:** Intermediate  
**Audience:** SharePoint Online Administrators, Microsoft 365 Administrators, IT Professionals  
**Prerequisites:** Basic Microsoft 365 administration and PowerShell knowledge

---

## Day 1 – Tenant Foundations & Site Management

---

### Module 1: Introduction to Microsoft 365 and SharePoint Online

#### Topics
- Microsoft 365 service architecture
- Role of SharePoint Online in Microsoft 365
- SharePoint Online vs SharePoint Server (conceptual comparison)
- Microsoft 365 Admin Center overview
- SharePoint Admin Center (modern)
- Service limits, quotas, and boundaries

#### Lab: Exploring the Microsoft 365 Environment
- Access Microsoft 365 Admin Center
- Review SharePoint Online tenant settings
- Check service health and message center

---

### Module 2: Identity, Access, and External Sharing

#### Topics
- Microsoft Entra ID fundamentals
- User and group identity models
- Admin roles vs site permissions
- Guest access and B2B collaboration
- External sharing policies (tenant and site)
- Conditional Access (overview)

#### Lab: Configuring Secure Access
- Configure external sharing policies
- Add and test guest users
- Assign SharePoint admin roles

---

### Module 3: Working with Site Collections

#### Topics
- Modern site collections (Team and Communication sites)
- Microsoft 365 Groups and Teams-connected sites
- Creating site collections (UI and PowerShell)
- Site ownership and security
- Storage quotas and site limits
- Site deletion and restore

#### Lab: Managing Site Collections
- Create modern site collections
- Configure storage quotas
- Delete and restore a site

---

## Day 2 – Information Architecture, Search & Customization

---

### Module 4: Permissions and Collaboration Model

#### Topics
- SharePoint permission inheritance
- SharePoint groups vs Microsoft 365 groups
- Sharing links and access scopes
- OneDrive vs SharePoint sharing behavior
- Best practices for secure collaboration

#### Lab: Designing a Permission Model
- Break permission inheritance
- Configure library-level permissions
- Test sharing scenarios

---

### Module 5: Managing Metadata and the Term Store

#### Topics
- Information architecture principles
- Managed metadata vs site columns
- Term Store hierarchy
- Term groups, term sets, and terms
- Delegated term management
- Applying metadata to libraries

#### Lab: Creating and Managing Metadata
- Create a term group and term set
- Assign managed metadata columns

---

### Module 6: Search in SharePoint Online and Microsoft Search

#### Topics
- Microsoft Search architecture
- SharePoint search vs Microsoft Search
- Managed properties (overview and limitations)
- Search schema in SharePoint Online
- Bookmarks, Q&A, and Acronyms
- Search verticals

#### Lab: Configuring Search Experience
- Create bookmarks and Q&A
- Test search across SharePoint and Microsoft 365

---

### Module 7: Apps and Customization in SharePoint Online

#### Topics
- SharePoint customization models
- SharePoint Framework (SPFx) overview
- Tenant App Catalog
- Deploying apps to SharePoint Online
- App permissions and governance

#### Lab: Managing Apps
- Configure the App Catalog
- Deploy an app to a SharePoint site

---

## Day 3 – Governance, Compliance & Automation

---

### Module 8: Content Governance and Compliance with Microsoft Purview

#### Topics
- Microsoft Purview overview
- Retention policies and retention labels
- Sensitivity labels
- Records management (modern approach)
- eDiscovery (Standard and Premium)
- Data Loss Prevention (DLP)

#### Lab: Implementing Compliance Controls
- Create retention labels
- Apply sensitivity labels
- Create an eDiscovery case

---

### Module 9: OneDrive for Business Administration

#### Topics
- OneDrive architecture
- Sharing and sync controls
- Storage policies
- Device access controls
- User lifecycle considerations

#### Lab: Configuring OneDrive Settings
- Configure OneDrive tenant options

---

### Module 10: Administration and Automation with PowerShell

#### Topics
- SharePoint Online Management Shell
- Microsoft Graph PowerShell
- Common administrative scripts
- Reporting and auditing via PowerShell

#### Lab: Automating SharePoint Administration
- Generate site and permission reports
- Perform bulk administrative tasks

---

### Module 11: Monitoring, Auditing, and Operational Best Practices

#### Topics
- Audit logs and activity monitoring
- Usage analytics
- Governance best practices
- Common SharePoint Online admin pitfalls

#### Lab: Operational Review
- Review audit logs
- Build a SharePoint governance checklist

---

### Module 12 (Optional): SharePoint Workflow Automation with Power Automate and Power Apps

> Optional module for classes that want hands-on workflow automation beyond administration and reporting.

#### Topics
- When to use Power Automate vs SharePoint built-in automation options
- SharePoint lists as workflow data sources (requests, approvals, status tracking)
- Power Automate basics for SharePoint admins
	- Triggers: item created/modified
	- Actions: approvals, notifications, updates
	- Connections and permissions (who owns the flow)
- Power Apps overview for list-based apps
	- Create an app from a SharePoint list
	- Form customization basics (data entry + validation concepts)
- Governance and operations
	- Environment and DLP considerations (admin perspective)
	- Support model: ownership, service accounts, documentation, lifecycle

#### Lab (Optional): Build a Simple Request Workflow
- Create a `NW-Pxx` request list (or reuse an existing scenario list)
- Create a Power Automate flow:
	- When an item is created → start an approval → update Status
- Create a basic Power Apps app from the list (or customize the form)
- Document:
	- required permissions
	- what happens on success/failure
	- ownership and change control

---

## Course Completion Outcomes

After completing this course, students will be able to:
- Administer SharePoint Online using modern tools
- Secure collaboration using Entra ID and Purview
- Design scalable site and metadata architectures
- Manage search and content discovery
- Automate administrative tasks with PowerShell

---

## Certification Alignment
- MS-102: Microsoft 365 Administrator
- SC-300: Identity and Access Administrator

