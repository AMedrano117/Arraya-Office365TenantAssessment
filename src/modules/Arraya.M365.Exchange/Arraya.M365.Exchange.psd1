@{
    RootModule        = 'Arraya.M365.Exchange.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'b7b72a93-1e4f-4a63-a9fd-f1184a2f1d56'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Get-AllExchangeMailboxDetails'
        'Get-AllPublicFolderDetails'
        'Get-AllRecipientDetails'
        'Get-ExchangeGroupDetails'
        'Get-ExchangeHybridConfiguration'
        'Get-MailFlowRulesandConnectors'
        'Get-SMTPRelayConfiguration'
        'Get-ThirdPartySpamFilteringConfig'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
