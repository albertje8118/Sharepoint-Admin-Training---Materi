# Module 2 Slides — Identity, Access, and External Sharing

> Format: slide title + bullets + speaker notes

## Slide 1 — Module goals
Bullets:
- Identity foundation: Microsoft Entra ID
- Roles vs permissions (tenant vs site)
- External sharing model (org + site)
- Guest users (B2B) basics
- Conditional Access overview

Speaker notes:
- Reinforce shared-tenant approach: tenant-wide changes are Trainer-only.

## Slide 2 — Scenario continuity and shared-tenant rules
Bullets:
- Scenario: Project Northwind Intranet Modernization
- Partner: Fabrikam (external)
- One tenant shared by trainer + 10 admin participants
- Use Participant ID P01–P10 and `NW-Pxx-...` naming

Speaker notes:
- Prevent collisions and avoid disrupting the tenant.

## Slide 3 — Identity layer: Entra ID
Bullets:
- Authentication and directory
- Member vs guest
- External collaboration settings
- Policies that affect access (Conditional Access)

Speaker notes:
- Many “SharePoint problems” are actually identity/policy problems.

## Slide 4 — Admin roles vs SharePoint permissions
Bullets:
- Admin roles: access to admin centers and tenant-level actions
- SharePoint permissions: access within a site
- Least privilege principle

Speaker notes:
- Site owners are not automatically tenant admins; tenant admins aren’t automatically site owners.

## Slide 5 — External sharing: organization level
Bullets:
- SharePoint admin center: Policies > Sharing
- Baseline for SharePoint + OneDrive
- OneDrive cannot be more permissive than SharePoint

Speaker notes:
- Avoid “Anyone” unless you deliberately accept the risk.

## Slide 6 — External sharing: site level
Bullets:
- Active sites > Site settings > More sharing settings
- Site can be same or more restrictive than org baseline
- Use site-level exceptions for controlled collaboration

Speaker notes:
- This is where participants can work safely (on practice sites).

## Slide 7 — Guest users (B2B) life cycle
Bullets:
- Invite guest (Entra)
- Guest redeems invitation
- Guest access depends on Entra + SharePoint configuration

Speaker notes:
- Invitation state matters: Pending vs Accepted.

## Slide 8 — Conditional Access (overview)
Bullets:
- Require MFA / compliant devices
- Block risky sign-ins
- Apply different requirements for guests

Speaker notes:
- Introduce now; deeper identity security belongs to identity-focused training.

## Slide 9 — Lab 2 preview
Bullets:
- Review org-level sharing (Trainer-only)
- Configure site-level sharing for practice site
- Verify guest account status (Trainer-provisioned)
- Review admin roles vs site permissions

Speaker notes:
- Outcome: choose the correct scope for each control.

## Slide 10 — Knowledge check
Bullets:
- Tenant role vs site role
- Org sharing vs site sharing
- Guest access failure causes
- Conditional Access purpose

Speaker notes:
- Keep answers concise; focus on scope and troubleshooting logic.
