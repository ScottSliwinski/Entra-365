# Entra-365

# Microsoft 365 & Hybrid Administration Toolkit

### Practical, GUI-based PowerShell tools built around the work administrators actually do.

## Project Overview

This project brings together a growing collection of graphical administration tools for **Microsoft 365, Microsoft Entra ID, Exchange Online, SharePoint Online, and hybrid Windows Active Directory environments**.

The goal is to fill practical workflow gaps in Microsoft’s administration interfaces: tasks that require extensive clicking, switching between portals, manually comparing information, or writing PowerShell commands to accomplish routine work.

These tools turn those processes into accessible, repeatable workflows with searchable results, clear previews, reporting, and controlled actions.

The project starts with individual scripts that solve specific administrative problems. The longer-term vision is a modular administration application that brings those tools together in one consistent interface.

## Why This Project Exists

Microsoft’s admin centers provide extensive functionality, but everyday administration often crosses the boundaries between services.

A single task might involve:

- Finding an account in on-premises Active Directory.
- Checking its synchronized identity in Entra ID.
- Reviewing Microsoft 365 licenses.
- Investigating Exchange mailbox permissions.
- Managing calendar or OneDrive access.
- Exporting records before making changes.

Completing that work can require several portals, multiple PowerShell connections, and repeated manual checks.

This project aims to reduce that overhead by organizing tools around the administrator’s task—for example, onboarding a user, reviewing delegates, or decommissioning an account.

**The objective is to complement Microsoft’s administration tools with focused workflows that are missing, fragmented, or difficult to perform efficiently.**

## Core Goals

### Make powerful administration accessible

Provide graphical interfaces for tasks that would otherwise require administrators to write or adapt PowerShell commands.

Common interface features include:

- UPN-based searches.
- Searchable lists and selectable results.
- Side-by-side comparisons where useful.
- Checkboxes for choosing accounts or actions.
- Progress indicators and readable status messages.
- Clear error reporting.
- CSV and text exports.

### Support hybrid environments

Build tools with the realities of synchronized environments in mind.

An account may appear in Microsoft 365 while its authoritative properties or group memberships are managed in Windows Active Directory. Tools should identify that distinction and perform changes through the appropriate service.

The aim is to help administrators understand both **where an object is managed** and **where a change needs to happen**.

### Reduce repetitive work

Make common administrative tasks faster and more consistent through reusable workflows, bulk operations, and reporting.

Where practical, administrators should be able to complete several related steps in one session without repeatedly reconnecting, restarting scripts, or copying information between tools.

### Make changes reviewable

Give administrators enough information to understand the scope of an operation before applying it.

Depending on the tool, this may include:

- Displaying current settings and permissions.
- Previewing proposed changes.
- Selecting individual objects.
- Confirming removal or other consequential actions.
- Running in test or preview mode.
- Recording results and failures.

### Build tools that others can learn from

Keep the underlying PowerShell understandable and maintainable.

Beyond providing working utilities, this repository aims to help other administrators understand Microsoft 365 automation, hybrid identity management, and the reasoning behind each workflow.

## Functional Areas

The following areas describe the intended scope of the repository. Individual tools will have their own implementation status, requirements, and limitations.

| Area | Intended capabilities |
|---|---|
| **User administration** | Search accounts, review identity information, and perform common cloud and hybrid management tasks. |
| **Onboarding** | Guide account creation, collect required information, check synchronization, and produce setup reports. |
| **Offboarding** | Coordinate access removal, session revocation, reporting, mailbox handling, and other selected decommissioning steps. |
| **Group management** | Review and export memberships, distinguish cloud-managed from synchronized groups, and copy selected memberships between users. |
| **Mailbox delegation** | View, add, update, and remove Full Access, Send As, and Send on Behalf permissions. |
| **Calendar permissions** | Review calendar access, resolve permission entries to useful identities, and manage delegates. |
| **Licensing** | Review assigned licenses and support licensing reports and administration workflows. |
| **MFA and account audits** | Review relevant MFA information and filter accounts by criteria such as domain, account age, or account type. |
| **SharePoint and OneDrive** | Support access reviews, delegation, and selected stale-account cleanup workflows. |
| **Reporting and cleanup** | Identify issues such as disabled delegates, orphaned permission entries, and accounts requiring review. |

## Design Principles

### Built for daily IT operations

Tools should solve concrete administrative problems and make frequent tasks easier to complete.

### Search first

Where practical, tools should allow quick, targeted searches without requiring a full tenant inventory before the administrator can begin.

### Configurable across organizations

Tenant domains, server names, and organization-specific settings should be configurable rather than embedded in the code.

### Clear identity handling

Interfaces and reports should display useful identifiers, including UPNs where available, so administrators can distinguish accounts with similar names.

### Consistent experience

As the collection grows, tools should move toward common patterns for authentication, searching, confirmations, progress reporting, exports, and logging.

### Modular development

Individual utilities should remain useful on their own while supporting the longer-term goal of a unified administration application.

## Technology Direction

The project is centered on **PowerShell 7.x**, with Windows graphical interfaces and the relevant Microsoft administration modules and APIs.

Depending on the tool, integrations may include:

- Microsoft Graph PowerShell.
- Exchange Online PowerShell.
- Windows Active Directory PowerShell.
- SharePoint Online administration tools.

Each script should document its supported PowerShell version, module dependencies, required permissions, and any Windows or hybrid-environment requirements.

## Long-Term Vision

The long-term goal is an **all-in-one, modular Microsoft 365 and hybrid administration command center**.

Administrators should be able to connect to the services they need, find a user, review related information, and launch focused management workflows from a consistent interface.

The repository will also serve as a collection of standalone tools for administrators who only need a particular function.

## Project Status

This is an evolving project developed from real administrative needs and refined through practical use and feedback.

Features, interfaces, and module coverage will change as tools are improved and consolidated. This README describes the project’s direction; each tool’s documentation should be checked for its current capabilities and supported scenarios.

## Contributions

Bug reports, workflow suggestions, documentation improvements, and code contributions are welcome.

Useful feedback includes:

- Administrative tasks that are difficult to complete in the existing portals.
- Hybrid scenarios that need better handling.
- Reproducible errors with relevant environment details.
- Improvements to usability, reporting, or performance.

Please remove credentials, tokens, personal information, and confidential tenant data before submitting examples or logs.

---

**Built by an administrator, for administrators—with the goal of making everyday Microsoft 365 and hybrid management simpler, clearer, and more repeatable.**
