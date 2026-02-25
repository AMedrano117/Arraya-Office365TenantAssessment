# Microsoft 365 Tenant Automation (PowerShell + Graph) - Azure DevOps

PowerShell and Microsoft Graph automation for Microsoft 365 tenant assessments, reporting, and controlled remediation.

## Purpose
This repository contains reusable modules, assessment scripts, remediation scripts, and prompt standards used by engineering teams to safely assess and manage customer Microsoft 365 tenants.

## Getting Started
1. Install PowerShell 7+
2. Run `tools/bootstrap-dev.ps1`
3. Review `prompts/skills.md` and `prompts/coding-standards.md`
4. Use `prompts/task-prompt-template.md` when generating scripts with Codex
5. Validate with `tools/invoke-scriptanalyzer.ps1` and `tools/run-pester.ps1`
