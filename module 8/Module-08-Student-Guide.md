# Module 8 — Content Governance and Compliance with Microsoft Purview (Student Guide)

## Learning objectives
By the end of this module, you will be able to:
1. Describe what Microsoft Purview is used for in Microsoft 365.
2. Explain retention labels, retention label policies, and where they apply.
3. Explain sensitivity labels, label publishing policies, and replication time.
4. Describe the modern eDiscovery workflow (cases, holds, searches, export).
5. Explain why DLP is deployed using simulation and pilots.

---

## 1) Microsoft Purview in a SharePoint admin context
Microsoft Purview is the Microsoft 365 governance and compliance portal where organizations configure controls such as:
- data lifecycle & records management (retention)
- information protection (sensitivity labels)
- eDiscovery
- data loss prevention (DLP)

Admin mindset:
- These features are designed to reduce risk and meet regulatory/business requirements.
- Many settings can have tenant-wide impact; treat them as governed changes.

---

## 2) Retention: policies vs labels (what admins should remember)
Retention is about what to **keep** and what to **delete**, and under what conditions.

### 2.1 Retention policies (broad)
- Retention policies typically apply to locations (SharePoint, OneDrive, Exchange, etc.) and can retain and/or delete content.
- Often used for baseline governance at scale.

### 2.2 Retention labels (item-level)
- Retention labels are applied at the **item level** (document/email).
- An item can have **only one retention label** applied at a time.

### 2.3 Publishing retention labels (retention label policies)
Creating a label doesn’t make it available to users.
- You publish labels using a **retention label policy** that defines locations and scope.

Operational note:
- When retention labels are published to SharePoint/OneDrive, they can take time to appear for users. Plan labs and rollouts with replication time in mind.

---

## 3) Sensitivity labels: classification + protection
Sensitivity labels help classify and (optionally) protect content.

### 3.1 Key concept: labels are published to users/groups
Unlike retention labels that are published to locations, sensitivity labels are published to **users and groups** using label publishing policies.

### 3.2 Replication time matters (SharePoint/OneDrive)
When you publish or change labels, allow time for them to replicate to the service.
Operational best practice:
- Pilot with a few users first.
- Validate behavior.
- Then expand scope to more users.

Admin governance note:
- For training, avoid encryption-focused complexity unless the tenant is configured and tested for it.

---

## 4) eDiscovery (modern experience): what it is and why admins care
eDiscovery helps organizations identify, preserve, and export content as evidence for legal/regulatory matters.

Key building blocks:
- **Case**: a container for an investigation.
- **Hold**: preserves content so it isn’t permanently deleted.
- **Search**: finds relevant content in locations.
- **Export / review**: produces deliverables for legal/compliance workflows.

2026 alignment note:
- Classic eDiscovery experiences have been retired (per Microsoft guidance). Use the modern eDiscovery experience in the Microsoft Purview portal.

---

## 5) DLP (Data Loss Prevention): safe deployment approach
DLP policies detect and (optionally) prevent risky actions, such as sharing sensitive content externally.

Admin best practices:
1. Start with clear policy intent (what are you protecting, from whom, and where?).
2. Scope locations carefully (SharePoint sites, OneDrive accounts, etc.).
3. Use **simulation mode** to understand impact.
4. Pilot with a small set of users/sites before broad rollout.

---

## References (official)
- Retention labels (create/publish/apply): https://learn.microsoft.com/en-us/purview/create-apply-retention-labels
- Retention overview: https://learn.microsoft.com/en-us/purview/retention
- Sensitivity labels (get started): https://learn.microsoft.com/en-us/purview/get-started-with-sensitivity-labels
- Enable sensitivity labels for SharePoint/OneDrive files: https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files
- eDiscovery (modern): https://learn.microsoft.com/en-us/purview/edisc-get-started
- Assign eDiscovery permissions: https://learn.microsoft.com/en-us/purview/edisc-permissions
- DLP create/deploy: https://learn.microsoft.com/en-us/purview/dlp-create-deploy-policy
