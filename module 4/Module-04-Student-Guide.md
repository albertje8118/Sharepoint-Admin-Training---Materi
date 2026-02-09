# Module 4 — Permissions and Collaboration Model

## Module objectives
By the end of this module, you will be able to:
- Explain permission inheritance and permission scopes in practical admin terms
- Describe SharePoint groups vs Microsoft 365 group-connected permissions
- Design a permission model for a realistic collaboration scenario
- Implement and validate site/library/folder permissions safely in a shared tenant
- Explain (at a practical level) how SharePoint and OneDrive sharing policies relate at the org and site level

---

## 4.1 Permission fundamentals (what admins actually troubleshoot)

### Key terms
- **Principal**: the identity you grant permissions to (user, SharePoint group, or Microsoft Entra group).
- **Permission level**: a named set of permissions (for example, **Full Control**, **Edit**, **Read**).
- **Inheritance**: a library/folder/file uses the same permissions as its parent.
- **Unique permissions / permission scope**: the object stops inheriting and has its own access control list.

### Why inheritance matters
Inheritance is SharePoint’s default because it makes administration predictable. When inheritance is broken repeatedly (many unique permission scopes), permissions become:
- harder to audit,
- harder to troubleshoot,
- easier to misconfigure,
- and potentially slower to evaluate at scale.

Microsoft explicitly recommends keeping unique permissions under control to avoid performance and manageability issues. (See References.)

---

## 4.2 SharePoint groups vs Microsoft 365 group-connected permissions

### SharePoint groups (site-scoped)
Most SharePoint sites have three default SharePoint groups:
- **Owners** (typically Full Control)
- **Members** (typically Edit)
- **Visitors** (typically Read)

These groups are convenient because you can:
- manage membership once,
- apply them at the site, library, or folder level,
- and avoid granting access to many individuals directly.

### Microsoft 365 group-connected team sites
For group-connected team sites, Microsoft 365 group membership drives access:
- Group **owners** become site owners
- Group **members** become site members

You can still use SharePoint groups (for example, Visitors) and can grant permissions directly in SharePoint, but you should be clear on the intended governance model.

Operational tip:
- If your site is connected to Teams (or is a Teams channel site), permissions may be managed through Teams and can be read-only in SharePoint.

---

## 4.3 Sharing links (collaboration) vs permissions (governance)

### What a sharing link “means”
Sharing links are another way to grant access, often at the file/folder level. Link types vary by configuration and policy. Common behaviors to understand:
- **People in your organization** link types are designed for internal sharing.
- **Specific people** links are designed for targeted access; recipients authenticate.
- **Anyone** links (anonymous) may be disabled or restricted by policy; use only when explicitly approved.

As an administrator, you should be able to:
- identify whether access is coming from group/site permissions or from a sharing link,
- review existing access (including links),
- remove access by removing permissions or deleting/revoking a link.

### SharePoint vs OneDrive external sharing at the policy level
At the organization level, SharePoint and OneDrive sharing settings are configured together (with OneDrive allowed to be equal or more restrictive). Site-level settings can be more restrictive than org-level, but not more permissive.

---

## 4.4 Guided demo (trainer-led): permission model design decisions

In the Northwind scenario, you will implement a safe permission model that:
- keeps broad access at the site level (simple collaboration),
- isolates sensitive content in a dedicated library,
- uses groups (not individuals) for manageability,
- minimizes the number of broken-inheritance objects.

The lab will implement the model inside each participant’s `NW-Pxx-ProjectSite` so changes are isolated in the shared tenant.

---

## Summary
You should now be able to:
- explain inheritance and unique scopes,
- choose between site-wide vs library/folder permissions,
- understand group-connected site permission implications,
- validate permissions using modern and classic (“Check Permissions”) views,
- and recognize when access is granted by a link vs a group.

---

## Knowledge check
1. Why is breaking inheritance repeatedly considered risky in SharePoint?
2. On a group-connected team site, what is the relationship between Microsoft 365 group membership and SharePoint site permissions?
3. When would you prefer using a SharePoint group instead of granting permissions to individuals?
4. What’s the difference between “People in your organization” links and “Specific people” links?
5. What is the first place you would check if a user reports “Access Denied” to a file: site permissions, file permissions, or sharing links—and why?

---

## References (official)
- Manage Permission Scopes in SharePoint: https://learn.microsoft.com/en-us/sharepoint/manage-permission-scope
- Sharing and permissions in the SharePoint modern experience: https://learn.microsoft.com/en-us/sharepoint/modern-experience-sharing-permissions
- Shareable links: Anyone, Specific people, and People in your organization: https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization
- Manage sharing settings for SharePoint and OneDrive: https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- Overview of external sharing in SharePoint and OneDrive: https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview
- Troubleshooting “Access Denied” (includes Check Permissions steps): https://learn.microsoft.com/en-us/troubleshoot/sharepoint/administration/access-denied-or-need-permission-error-sharepoint-online-or-onedrive-for-business
