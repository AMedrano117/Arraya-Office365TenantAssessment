# Requirements
- PowerShell 7+
- Pester
- PSScriptAnalyzer
- Microsoft Graph PowerShell SDK modules as needed
- ExchangeOnlineManagement module
- MicrosoftTeams module (for scripts that collect Teams service details)
- ImportExcel module

Notes:
- `Office365Custom` version `1.2.1` is vendored in this repository under `src/vendor/Office365Custom/1.2.1`.
- Module import is handled by script/module entry points; users do not need to import it manually.
