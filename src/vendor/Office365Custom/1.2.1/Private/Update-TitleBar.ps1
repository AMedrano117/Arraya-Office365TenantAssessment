# Helper function to update the title bar
function Update-TitleBar {
    param (
        [string]$Title
    )
    $host.ui.RawUI.WindowTitle = $Title
}
