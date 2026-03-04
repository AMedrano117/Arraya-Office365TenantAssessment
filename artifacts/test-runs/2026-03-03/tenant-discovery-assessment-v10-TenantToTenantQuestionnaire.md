# Microsoft 365 Tenant to Tenant Questionnaire

**Date:** 2026-03-03
**Prepared By:** amedrano (automated assessment)
**Client:** Cavehouse Brewery
---

## Table of Contents

- [Tenant](#tenant)
- [Identity](#identity)
- [Email](#email)
- [Files](#files)
- [Communication](#communication)
- [Security](#security)
- [Compliance](#compliance)
- [Devices](#devices)
- [Additional Workloads](#additional-workloads)
- [External Resources](#external-resources)
- [Additional Comments or Concerns](#additional-comments-or-concerns)

---

## Tenant

| Discovery Item | Answer |
|---------------|--------|
| What licensing SKUs exist in the tenant and how many of each? | CCI Bots Private Preview: 3/10000; EXCHANGEENTERPRISE: 0/3; MCOCAP: 1/1; MCOPSTN1: 1/1; Microsoft Power Apps Plan 2 Trial: 1/10000; Microsoft Power Automate Free: 12/10000; Microsoft_365_ Business_ Premium_(no Teams): 4/12; Microsoft_365_Copilot: 5/6; Microsoft_Teams_Enterprise_New: 4/15; Office 365 E3: 1/1; Power BI Premium Per User: 2/16; POWERAPPS_PER_USER: 1/1; SPE_E5: 11/11; Windows Store for Business: 0/25 |
| What is your tenant address? | Default domain: cavehousebrewing.com; Initial domain: cavehousebrew.onmicrosoft.com |
| Is your tenant homed in the United States? If not, where? | Yes. Tenant country/data location indicates United States/North America. |
| Is your tenant a Multi-Geo environment? If so, where is your central location? | No multi-geo configuration detected in tenant overview data. |
| Does your organization utilize Microsoft Copilot? | Likely yes for Microsoft Copilot; matched license(s): Microsoft_365_Copilot (5/6 consumed) |
| Does your organization utilize Copilot Studio? | Likely yes for Copilot Studio; matched license(s): CCI Bots Private Preview (3/10000 consumed); Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Microsoft_365_Copilot (5/6 consumed); Office 365 E3 (1/1 consumed) |
| Do your users utilize Microsoft Power Automate? | Likely yes for Power Automate; matched license(s): CCI Bots Private Preview (3/10000 consumed); Microsoft Power Apps Plan 2 Trial (1/10000 consumed); Microsoft Power Automate Free (12/10000 consumed); Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed) |
| Do your users utilize Microsoft Power Apps? | Likely yes for Power Apps; matched license(s): Microsoft Power Apps Plan 2 Trial (1/10000 consumed); Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); POWERAPPS_PER_USER (1/1 consumed) |
| Does your organization use Power Pages? | Likely yes for Power Pages; matched license(s): POWERAPPS_PER_USER (1/1 consumed) |
| Does your organization use Dynamics 365? | Not detected in discovered license inventory. |
| Does your organization use Power BI? | Likely yes for Power BI; matched license(s): Power BI Premium Per User (2/16 consumed) |
| Do you have self-service purchases enabled? | Unknown; MSCommerce module install/load failed: Failure from remote command: Import-Module -Name 'MSCommerce': The specified module 'MSCommerce' was not loaded because no valid module file was found in any module directory. |
| Does the organization utilize the source Entra tenant for Azure resources? | Not collected; Azure module not used in this report |

---

## Identity

> **Note:** Entra ID is the new name for Azure Active Directory (Azure AD)

| Discovery Item | Answer |
|---------------|--------|
| What identity provider is used for sign-on? | Entra ID cloud authentication appears to be the primary sign-on provider. |
| Are objects synchronized from on-prem Active Directory? | Yes. Directory synchronization is enabled. Last reported sync: 2023-11-02. |
| Is self-service password reset or password writeback enabled? | Not collected by current script; manual validation required. |
| What MFA solution is used? | Entra ID MFA. Methods: microsoftAuthenticatorAuthenticationMethodConfiguration, smsAuthenticationMethodConfiguration, softwareOathAuthenticationMethodConfiguration, voiceAuthenticationMethodConfiguration, emailAuthenticationMethodConfiguration; CA policies requiring MFA: 0; registered users: 30.6%. |
| Is Entra ID Enterprise SSO used for SaaS applications? | No enterprise SSO applications were detected in the current run. |
| If yes, provide a list of applications | No SSO application list collected. |
| Is Entra ID Conditional Access used? | Yes. 10 Conditional Access policy/policies collected. |
| Is Entra ID App Proxy used? | Not collected by current script; manual validation required. |
| Is Entra ID Private Access used? | Not collected by current script; manual validation required. |
| Is Privileged Identity Management (PIM) used? | Not collected by current script; manual validation required. |

---

## Email

| Discovery Item | Answer |
|---------------|--------|
| Active user mailboxes to migrate | 22 active user mailbox(es). |
| Inactive user mailboxes to migrate | 14 inactive mailbox(es). |
| Shared mailboxes to migrate | 266 shared mailbox(es). |
| Resource mailboxes to migrate | 0 resource mailbox(es). |
| Archive mailboxes to migrate | 3 archive-enabled mailbox(es). |
| Distribution lists to migrate | 1 distribution list(s). |
| Dynamic distribution lists | 0 dynamic distribution group(s). |
| External contacts to migrate | 0 external contact/mail user object(s). |
| Email domains to migrate | 4 custom domain(s): lvhndemo.cavehousebrewing.com, lab.com, jeffersondemo.cavehousebrewing.com, cavehousebrewing.com |
| Public folders | 2 public folder object(s). |
| Average mailbox size | 0 GB average primary mailbox size. |
| Largest mailbox size | 0.65 GB (Michael Wishnefsky). |
| Hybrid Exchange configuration | Possible Hybrid; type: Mail Flow (Single On-Premises connector); evidence: OrganizationRelationship configured (1); Inbound OnPremises connector(s) (1); Mail flow On-Premises connector(s) (1). |
| On‑prem Exchange servers | Not collected by current script; manual validation required. |
| Edge Transport servers | Not collected by current script; manual validation required. |
| Mail transport rules | 5 transport rule(s) collected. |
| Exchange Online Message Protection / Encryption | Microsoft-native protection/encryption capabilities are indicated by licensing, but specific OME or mail-flow encryption configuration was not collected. |
| PST usage | Not collected by current script; manual validation required. |
| 3rd‑party mail hygiene | Not detected from current mail flow analysis. Inbound connectors: 1, outbound connectors: 1. |
| 3rd‑party archiving | Not collected by current script; manual validation required. |
| Additional address lists | Not collected by current script; manual validation required. |
| Send/receive connectors | 1 inbound and 1 outbound connector(s) |

---

## Files

| Discovery Item | Answer |
|---------------|--------|
| OneDrive storage in use | 290748.2 GB across 31 OneDrive site(s). |
| SharePoint Online storage in use | 2655914.68 GB across 108 SharePoint site(s). |
| External sharing enabled | Not collected by current script; manual validation required. |
| Data sync restricted to managed devices | Not collected by current script; manual validation required. |
| SharePoint sites to migrate | 108 SharePoint site(s). |
| All sites modern | Undetermined from current site inventory alone; manual validation recommended. |
| User-created SharePoint sites allowed | Not collected by current script; manual validation required. |
| SharePoint hybrid configuration | Not collected by current script; manual validation required. |
| Active OneDrive profiles | 31 active OneDrive profile(s). |
| Archived OneDrive profiles | 0 archived OneDrive profile(s). |
| SharePoint workflows | Not collected by current script; manual validation required. |

---

## Communication

| Discovery Item | Answer |
|---------------|--------|
| Microsoft Teams to migrate | Teams topology was not collected in the current PowerShell 7 app-only run. Microsoft 365 groups detected: 34. Manual validation required. |
| Non‑Teams Microsoft 365 Groups | 34 Microsoft 365 group(s) detected in Entra. Team association is incomplete in this run, so manual validation is required. |
| Users can create Teams | Not collected by current script; manual validation required. |
| External guest access enabled | Yes or likely yes. Guest/external users detected: 11; cross-tenant partner relationships: 3. |
| Teams Audio Conferencing | Likely yes for Teams Audio Conferencing; matched license(s): SPE_E5 (11/11 consumed) |
| Teams PSTN calling | Voice users: 10; data source: GraphLicenseInference; notes: Teams PowerShell is not connected. Voice-enabled users are inferred from assigned voice licenses and enabled service plans.. |
| Teams Rooms | Not detected in discovered license inventory. |
| Teams room hardware | Not collected by current script; manual validation required. |

---

## Security

| Discovery Item | Answer |
|---------------|--------|
| Microsoft Defender products in use | Likely yes for Microsoft Defender; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); SPE_E5 (11/11 consumed) |
| Non‑Outlook client access allowed | Not collected by current script; manual validation required. |
| MFA usage scenarios | Registered users: 33/108; methods enabled: microsoftAuthenticatorAuthenticationMethodConfiguration, smsAuthenticationMethodConfiguration, softwareOathAuthenticationMethodConfiguration, voiceAuthenticationMethodConfiguration, emailAuthenticationMethodConfiguration. |
| Insider Risk Management | Not collected by current script; manual validation required. |
| Customer Lockbox | Not collected by current script; manual validation required. |
| Background checks required for Arraya access | Customer policy item; manual validation required. |

---

## Compliance

| Discovery Item | Answer |
|---------------|--------|
| eDiscovery solution | Microsoft Purview/eDiscovery capability appears licensed. Active case inventory was not collected in this run. |
| Active eDiscovery cases | Not collected by current script; manual validation required. |
| Data labeling & classification | Labeling and classification capability is indicated by licensing, but policy inventory was not collected. |
| Automatic classification | Not collected by current script; manual validation required. |
| Encryption of data at rest | Microsoft 365 encrypts customer data at rest by platform default. Customer-managed key usage was not collected in this run. |
| Retention period | Not collected by current script; manual validation required. |
| Regulatory requirements (HIPAA, ITAR, etc.) | Not collected by current script; manual validation required. |
| Data Loss Prevention (DLP) | Potential DLP capability is indicated by licensing, but actual DLP policy inventory was not collected. |
| Endpoint DLP | Potential endpoint DLP capability is indicated by licensing, but endpoint DLP policy inventory was not collected. |
| Other Purview capabilities | Detected Purview-related licensing. Workbook and JSON outputs should be used for detailed follow-up. |
| Encryption key type | Not collected by current script. Default assumption is Microsoft-managed keys unless Customer Key has been separately configured. |

---

## Devices

| Discovery Item | Answer |
|---------------|--------|
| Microsoft Intune usage | Likely yes. Managed devices: 9; devices reporting MDM values of Intune or SCCM, Not Managed: 9. |
| Device configuration policies | Not collected by current script; manual validation required. |
| Device compliance policies | Not collected by current script; manual validation required. |
| Defender for Endpoint via Intune | Defender licensing is present, but Defender for Endpoint deployment or configuration via Intune was not collected. |
| On‑prem Endpoint Manager (SCCM) | Possible. 9 device(s) reported MDM values of Intune or SCCM, Not Managed, but SCCM-specific inventory was not collected. |
| Applications deployed via Intune | Not collected by current script; manual validation required. |
| Autopilot usage | Not collected by current script; manual validation required. |
| Apple Business Manager / Android Enterprise | Not collected by current script; manual validation required. |
| Co‑management enabled | Not collected by current script; manual validation required. |
| Supported workstations | 39 Windows and 0 macOS workstation device(s) discovered. |
| Corporate mobile devices | Mobile devices discovered: 2 (iOS: 2, Android: 0). Ownership or corporate classification was not collected. |
| BYOD mobile devices | Not collected by current script; manual validation required. |
| Employee‑owned device management | Not collected by current script; manual validation required. |
| iOS vs Android percentages | iOS: 100%; Android: 0%. |
| Device join type (Hybrid/Cloud) | Trust or join types observed: AzureAd, ServerAd, Workplace. |
| Standard browser | Not collected by current script; manual validation required. |

---

## Additional Workloads

| Discovery Item | Answer |
|---------------|--------|
| Microsoft Viva | Likely yes for Microsoft Viva; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |
| Microsoft Forms | Likely yes for Microsoft Forms; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Microsoft_Teams_Enterprise_New (4/15 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |
| Microsoft Planner | Likely yes for Microsoft Planner; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |
| Microsoft Yammer | Likely yes for Microsoft Yammer/Viva Engage; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |
| Microsoft Stream | Likely yes for Microsoft Stream; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |
| Project / Visio licensing | Likely yes for Project/Visio; matched license(s): Microsoft_365_ Business_ Premium_(no Teams) (4/12 consumed); Office 365 E3 (1/1 consumed); SPE_E5 (11/11 consumed) |

---

## External Resources

| Discovery Item | Answer |
|---------------|--------|
| 3rd‑party email services | Review required. Connectors: 2; remote domains: 2; third-party filtering detected: No. |
| 3rd‑party device management | Current device inventory reports these MDM values: Intune or SCCM, Not Managed. |
| 3rd‑party cloud storage | Not collected by current script; manual validation required. |

---

## Additional Comments or Concerns

- Automated tenant discovery generated on 2026-03-03 for Cavehouse Brewery.
- Primary migration review items: Cross-tenant and B2B settings (Review); Public folders (Review); Verified custom domains (Blocker); Mail routing (Review); Directory synchronization (Review); Teams workload data (Needs Data).
- Current assessment findings to review: License 'SPE_E5' has 0 remaining (fully allocated) Domain 'lab.com' is not verified License 'POWERAPPS_PER_USER' has 0 remaining (fully allocated) License 'Office 365 E3' has 0 remaining (fully allocated)
- Manual validation is still required for Teams topology details, SharePoint external sharing posture, App Proxy/Private Access/PIM, compliance policy inventory, PST usage, and browser standards.


---

---

**Confidential**  
Version 3
