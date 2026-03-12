function Get-ArrayaAssessmentRecommendationText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Finding
    )

    $anchor = [string]$Finding.Anchor
    switch ($anchor) {
        'licenses' { return 'Review paid SKU utilization and purchase additional licenses before assignment failures occur.' }
        'domains' { return 'Verify domain ownership and DNS posture before migration cutover.' }
        'domains-dns' { return 'Harden SPF/DKIM/DMARC and review anti-spoofing exceptions to align with Microsoft mail-auth guidance.' }
        'identity-admins' {
            switch ([string]$Finding.Category) {
                'Inactive Admin Accounts' { return 'Review privileged accounts for inactivity, disable unused admin identities, and maintain emergency access coverage.' }
                'Guest Access' { return 'Review guest and external identities for lifecycle governance and least-privilege access before migration.' }
                'Unmanaged Groups' { return 'Assign group owners and verify owner health for collaboration and access-governance continuity.' }
                'Admin Sign-in Telemetry' { return 'Ensure sign-in telemetry is available for privileged accounts (app permissions/log retention) so stale-admin monitoring is reliable.' }
                'Emergency Access Accounts' { return 'Maintain at least two cloud-only emergency access accounts, monitor them, and keep them excluded from daily operational use.' }
                'Emergency Access CA Exclusions' { return 'Validate Conditional Access emergency-access exclusions so break-glass accounts can sign in during policy or identity outages.' }
                default { return 'Validate admin access, guest usage, and group ownership before identity migration activities.' }
            }
        }
        'mailboxes' { return 'Identify oversized or specialized mailboxes early to plan batching, archives, and exception handling.' }
        'compliance-retention' { return 'Review oversized mailboxes and archives for litigation/retention hold coverage and remediate objects without explicit hold controls.' }
        'inactive-mailboxes' { return 'Decide whether inactive mailboxes need retention, restore, or exclusion from scope.' }
        'sharepoint-onedrive' { return 'Use site inventory, ownership, and storage metrics to prioritize high-risk collaboration workloads.' }
        'devices' { return 'Review stale and non-compliant devices before identity and endpoint cutover.' }
        'ad-connect' { return 'Document synchronization dependencies and plan cloud identity cutover or staged decommissioning.' }
        'conditional-access-mfa' {
            switch ([string]$Finding.Category) {
                'Legacy Authentication' { return 'Block legacy authentication with Conditional Access and verify modern-auth readiness before enforcement.' }
                'Risk-based Conditional Access' { return 'Implement sign-in risk and user risk Conditional Access policies to align with Microsoft identity protection practices.' }
                'Admin Consent Workflow' { return 'Enable and tune the admin consent request workflow so app consent escalations follow governance controls.' }
                'App Consent Governance' { return 'Harden app consent and app registration settings to reduce over-privileged or unmanaged enterprise app risk.' }
                default { return 'Review CA and MFA design to avoid post-migration lockouts or authentication regressions.' }
            }
        }
        'exchange-hybrid' { return 'Validate hybrid, connectors, and migration endpoints because they affect tenant-to-tenant messaging strategy.' }
        'cross-tenant-access' { return 'Review cross-tenant and B2B settings for coexistence, external collaboration, and post-migration cleanup.' }
        'secure-score' { return 'Use the mapped Microsoft Secure Score action to prioritize remediation with the highest security impact.' }
        'ownership-governance' {
            switch ([string]$Finding.Category) {
                'Unowned Objects' { return 'Assign at least one accountable owner to each collaboration object and validate ownership handoff before migration or governance workflows.' }
                'Owner Health' { return 'Reassign ownership from disabled or stale accounts to active custodians and formalize backup ownership coverage.' }
                'OneDrive Ownership Mismatch' { return 'Review OneDrive sites where the current owner differs from the URL-derived default user and confirm documented stewardship.' }
                default { return 'Review ownership governance tables and assign healthy owners for all unmanaged or mismatched objects.' }
            }
        }
        'employee-experience-insights' {
            switch ([string]$Finding.Category) {
                'Report Identity Visibility' { return 'Adjust Microsoft 365 usage report settings so actor-level names are available to authorized analysts when needed for operations and investigations.' }
                'Usage Telemetry Coverage' { return 'Confirm reports permissions and data-retention windows so usage telemetry is complete for business and adoption analysis.' }
                'Collaboration Activity Baseline' { return 'Track high-volume collaboration patterns and low-adoption areas to prioritize enablement and governance workstreams.' }
                default { return 'Use collaboration activity telemetry to guide adoption, governance, and operational readiness improvements.' }
            }
        }
        default { return 'Review the related worksheet and validate whether remediation is required for your migration or security objectives.' }
    }
}
