<# 
.SYNOPSIS
   Loads a hashtable from a JSON file.
.DESCRIPTION
   This function reads a JSON file from the specified path, converts the JSON content into an object,
   and then builds a hashtable by filtering out any properties with blank names or null values.
.PARAMETER FilePath
   The full file path to the JSON file containing the hashtable.
.EXAMPLE
   $myHash = Load-HashTableFromJson -FilePath "C:\Temp\myhash.json"
#>
function Load-HashTableFromJson {
    param (
        [string]$FilePath
    )

    if (Test-Path -Path $FilePath) {
        try {
            $json = Get-Content -Path $FilePath -Raw
            $jsonObject = ConvertFrom-Json -InputObject $json

            # Create a new hashtable and clean up the JSON object
            $HashTable = @{}
            Write-Verbose "Cleaning up JSON object of empty or null values."
            foreach ($prop in $jsonObject.PSObject.Properties) {
                if (![string]::IsNullOrWhiteSpace($prop.Name) -and $prop.Value -ne $null) {
                    $HashTable[$prop.Name] = $prop.Value
                }
            }

            Write-Host "Hash table loaded from $FilePath" -ForegroundColor Green
            return $HashTable
        } catch {
            Write-Host "Error loading hash table from $($FilePath): $($_.Exception.Message)" -ForegroundColor Red
            return @{}
        }
    } else {
        Write-Host "File not found: $FilePath" -ForegroundColor Red
        return @{}
    }
}
