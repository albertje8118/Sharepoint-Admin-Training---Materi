# Module 2: Identity, Access, and External Sharing

## Course lab scenario (continuity)
This course uses one shared Microsoft 365 tenant (trainer + 10 admin participants) and one scenario story.

- Scenario overview: ../scenario/Lab-Scenario-Overview.md

In this module:
- We treat **tenant-wide policy** changes as **Trainer-only** unless your trainer explicitly asks you to change them.
- Participant hands-on work is designed to be safe when applied to participant-isolated practice sites.

---

## Module objectives
After completing this module, you will be able to:
- Explain Microsoft Entra ID fundamentals relevant to SharePoint Online.
- Distinguish admin roles (tenant) from SharePoint permissions (site).
- Describe guest access (B2B collaboration) and how it interacts with SharePoint sharing.
- Identify and apply the correct external sharing control at the correct scope (org vs site).
- Explain Conditional Access at a high level.

---

## 1. Microsoft Entra ID fundamentals (for SharePoint admins)

### Identity and access in Microsoft 365
SharePoint Online authentication and authorization are rooted in Microsoft Entra ID:
- Users sign in and authenticate via Entra ID.
- Access is evaluated based on:
  - User identity (member vs guest)
  - Role assignments (admin roles)
  - Policies (Conditional Access, external collaboration restrictions)
  - Resource configuration (SharePoint tenant and site sharing settings)

### Member vs guest (B2B collaboration)
In a typical workforce tenant:
- **Members** are internal users.
- **Guests** are external users invited into your directory for collaboration.

Guest user behavior is impacted by:
- Entra external collaboration settings (who can invite guests, restrictions)
- SharePoint/OneDrive sharing configuration
- Site-level sharing controls

---

## 2. Admin roles vs SharePoint permissions (don’t confuse these)

### Tenant admin roles (directory / Microsoft 365 level)
Admin roles such as **SharePoint Administrator** and **Global Administrator** provide access to admin centers and the ability to manage tenant-level configuration.

Key idea:
- Admin roles are **not the same thing** as being a site owner.

### SharePoint permissions (resource level)
SharePoint permissions control what users can do inside a specific site/library/list.

Examples:
- Site Owner: typically full control on a site, but not necessarily a tenant admin.
- Site Member/Visitor: collaboration/consumption roles.

Practical takeaway:
- Use **least privilege**: give people only what they need (admin role vs site role).

---

## 3. External sharing: the control stack

External sharing decisions are made through multiple layers:

1) **Entra external collaboration settings**
- Who is allowed to invite guests?
- Are there restrictions on invitations?

2) **Organization-level sharing** (SharePoint admin center)
- A tenant-wide baseline for SharePoint and OneDrive external sharing.
- Site-level sharing can be the same or more restrictive, not more permissive.

3) **Site-level sharing** (SharePoint admin center)
- A targeted control for a specific site.

### Safe operational pattern
- Start by understanding the organization-level baseline.
- Apply exceptions at the site level (where appropriate), ideally for an isolated extranet or project site.

---

## 4. Guest access and common admin pitfalls

Common causes of “guest can’t access” issues:
- Org-level sharing is too restrictive.
- Site-level sharing is more restrictive than expected.
- Guest was invited but did not redeem the invitation.
- The resource is connected to a Microsoft 365 Group/Teams and group guest settings affect access.
- Conditional Access or device restrictions block access.

---

## 5. Conditional Access (overview)

Conditional Access is an identity-driven policy layer that can require or block access based on conditions such as:
- User risk or sign-in risk
- Device compliance
- Location/network
- Authentication strength

In SharePoint scenarios, Conditional Access often explains:
- Why users can sign into Microsoft 365 but cannot access SharePoint resources.
- Why guests behave differently from members.

---

## Summary
In this module you learned how identity and policy layers (Entra ID) intersect with SharePoint tenant settings and site-level sharing configuration, and how to apply least-privilege thinking to roles and permissions.

---

## Knowledge check (self-assessment)
1) What is the difference between a SharePoint Administrator role and a site owner?  
2) Why can a site-level sharing setting never be more permissive than the organization-level sharing setting?  
3) Name two common causes of guest access failures.  
4) What is Conditional Access trying to accomplish?

### Suggested answers
1) Admin role grants admin center/tenant capabilities; site owner grants permissions within a specific site.  
2) The organization-level setting is the tenant baseline; sites inherit that maximum permissiveness.  
3) Example: org sharing disabled; site sharing too restrictive; invitation not redeemed; CA restrictions.  
4) Enforce access requirements/controls based on identity and context.

---

## References (Microsoft Learn)
- Manage sharing settings for SharePoint and OneDrive (org-level): https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off
- Change the sharing settings for a site (site-level): https://learn.microsoft.com/en-us/sharepoint/change-external-sharing-site
- Add B2B guest users (Entra): https://learn.microsoft.com/en-us/entra/external-id/add-users-administrator
- Assign admin roles (Microsoft 365 admin center): https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/assign-admin-roles
- About the SharePoint Administrator role: https://learn.microsoft.com/en-us/sharepoint/sharepoint-admin-role
