# coding-standards.md
## PowerShell + Microsoft Graph Coding Standards
- Prefer PowerShell 7+
- Use advanced functions and typed/validated parameters
- Use try/catch/finally and actionable errors
- Prefer Graph SDK, use Invoke-MgGraphRequest when SDK is missing support
- Explicitly state v1.0 vs beta
- Handle pagination and throttling retries
- Export CSV/JSON outputs consistently
- No hardcoded secrets/credentials/tokens
- Document prerequisites, scopes, roles, assumptions, examples, outputs, limitations
