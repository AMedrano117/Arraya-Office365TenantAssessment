<# 
.SYNOPSIS
   Saves a hashtable to a JSON file.
.DESCRIPTION
   This cmdlet cleans the provided hashtable by removing any keys with empty or null values,
   converts the cleaned hashtable to JSON (with a depth of 10), and writes it to a specified file path.
.PARAMETER HashTable
   The hashtable to clean and convert to JSON.
.PARAMETER FilePath
   The full file path where the JSON content will be saved.
.EXAMPLE
   Save-HashTableToJson -HashTable $myHash -FilePath "C:\Temp\myhash.json"
#>
function Save-HashTableToJson {
    param (
        [hashtable]$HashTable,
        [string]$FilePath
    )

    # Validate and clean up the hash table
    $cleanedHashTable = @{}
    Write-Verbose "Cleaning up hash table of empty or null values."
    foreach ($key in $HashTable.Keys) {
        if (![string]::IsNullOrWhiteSpace($key) -and $HashTable[$key] -ne $null) {
            $cleanedHashTable[$key] = $HashTable[$key]
        }
    }

    $json = $cleanedHashTable | ConvertTo-Json -Depth 10
    Set-Content -Path $FilePath -Value $json
    Write-Host "Hash table saved to $FilePath" -ForegroundColor Green
}
