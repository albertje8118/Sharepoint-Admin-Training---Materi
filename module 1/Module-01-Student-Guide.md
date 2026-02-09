# Module 1: Introduction to Microsoft 365 and SharePoint Online

## Course lab scenario (continuity)
This course uses one shared Microsoft 365 tenant for the whole class (trainer + 10 admin participants) and a single scenario story.

- Scenario overview: ../scenario/Lab-Scenario-Overview.md

In Module 1, labs focus on **orientation and verification**. In later modules, configuration work will be scoped either to **trainer-only tenant-wide changes** or to **participant-isolated practice sites**.

## Module objectives
After completing this module, you will be able to:
- Describe Microsoft 365 service architecture at a practical admin level.
- Explain the role of SharePoint Online in Microsoft 365 collaboration and content services.
- Compare SharePoint Online and SharePoint Server conceptually (deployment, lifecycle, governance model).
- Navigate Microsoft 365 admin center and SharePoint admin center to perform baseline tenant checks.
- Locate and interpret service limits, quotas, and boundaries using official Microsoft documentation.

---

## 1. Microsoft 365 service architecture (admin mental model)

### What Microsoft 365 is (from an admin perspective)
Microsoft 365 is a set of cloud services (“workloads”) backed by:
- **Identity** (Microsoft Entra ID): authentication, authorization, device and conditional access controls
- **Workloads** (SharePoint, Exchange, Teams, OneDrive, Purview, etc.)
- **Admin surfaces** (Microsoft 365 admin center, workload-specific admin centers)
- **APIs & automation** (Microsoft Graph, workload PowerShell modules)

### Why this matters to SharePoint admins
Many SharePoint outcomes are influenced by tenant-wide settings outside SharePoint, such as:
- Identity and access policies (Entra ID, Conditional Access)
- Compliance controls (Microsoft Purview)
- Search and content experiences integrated across Microsoft 365

---

## 2. Role of SharePoint Online in Microsoft 365

### SharePoint Online as a service
SharePoint Online (SPO) provides:
- **Sites** for team collaboration, communication, and intranet publishing
- **Document libraries** for controlled storage, sharing, versioning, metadata
- **Content services** that integrate with:
  - Microsoft Teams (Teams-connected SharePoint sites)
  - OneDrive (personal storage built on SharePoint technology)
  - Microsoft Search (content discovery across Microsoft 365)
  - Power Platform (content sources and connectors)

### Common admin responsibilities
As a SharePoint Online administrator you typically manage:
- Tenant-level policies (sharing, access, site creation governance)
- Site lifecycle (creation, ownership, storage, deletion/restore)
- Information architecture building blocks (metadata/term store—covered later)
- Operational health (service health, message center, admin notifications)
- Governance and compliance alignment (Purview policies, audit visibility)

---

## 3. SharePoint Online vs SharePoint Server (conceptual comparison)

### How to compare (recommended framing)
Avoid a “feature checklist” mentality. Compare based on operating model:

1) **Deployment & ownership**
- SharePoint Server: your infrastructure, patching, upgrades, capacity planning
- SharePoint Online: Microsoft operates the service; you control configuration and governance

2) **Change cadence**
- Server: changes follow your maintenance and upgrade schedule
- Online: continuous service updates; admin controls focus on governance and adoption readiness

3) **Integration**
- Server: integrations often require custom design and on-prem dependencies
- Online: designed for Microsoft 365 integration (Entra ID, Purview, Teams, Search, Graph)

4) **Customization approach**
- Server: historically more farm/solution-level customization patterns
- Online: modern customization favors client-side solutions and API-based automation (SPFx/Graph patterns)

### Admin takeaway
SPO administration is primarily about **policy, governance, lifecycle, and integration** rather than server operations.

---

## 4. Microsoft 365 admin center overview (what SharePoint admins should use it for)

### Typical tasks for SharePoint admins
- Confirm tenant context and admin roles
- Review **Service health** for incidents/advisories affecting SharePoint/OneDrive
- Review **Message center** for upcoming changes that may impact SharePoint governance
- Validate user licensing assignments when troubleshooting access

### Good practice
Build a habit of checking:
- Service health first (is it you or is it the service?)
- Message center next (is there a change rolling out that explains the behavior?)

---

## 5. SharePoint admin center (modern) overview

### What it’s for
SharePoint admin center focuses on:
- Tenant sharing and access policies (high-level)
- Site management and lifecycle
- Storage management and reports
- Org-wide settings affecting SharePoint and (often) OneDrive

### Admin navigation principle
UI labels can change between tenants and over time. Your goal is to understand:
- Which settings are tenant-wide vs site-specific
- Which settings are “policy” vs “operational configuration”
- How to validate the effective setting on a specific site

---

## 6. Service limits, quotas, and boundaries

### What these terms mean in practice
- **Limits/boundaries:** hard platform constraints (what the service can support)
- **Quotas:** configurable allocations (e.g., storage allocation model), where applicable
- **Recommendations:** guidance for performance and manageability

### How to manage this safely in real work
- Do not rely on memorized numbers.
- Always verify current values in Microsoft Learn (limits are updated over time).
- Document which limit matters to your scenario (sites, storage, list/library behaviors, sharing constraints, etc.).

### Exercise (no tools required)
Pick one scenario and write down which limit category you would research:
- “We need 200k items in a library”
- “We need external sharing to vendors for a project site”
- “We need an intranet hub with many associated sites”
- “We want to retain content for 7 years”

---

## Summary
In this module you learned how SharePoint Online fits in Microsoft 365, how the admin experience is structured across admin centers, and how to approach limits and boundaries using authoritative documentation rather than tribal knowledge.

---

## Knowledge check (self-assessment)
1) Which Microsoft 365 component is responsible for identity and access control?  
2) Why should a SharePoint admin care about Message center?  
3) Name two conceptual differences between SharePoint Online and SharePoint Server.  
4) What is the safe approach to “service limits” in documentation?

### Suggested answers
1) Microsoft Entra ID.  
2) It announces changes that can affect governance, features, and user experience.  
3) Example: cloud-operated vs self-operated; continuous updates vs scheduled upgrades; integration model differs.  
4) Verify current values in official docs; avoid relying on memorized numbers; document scenario-specific limits.
