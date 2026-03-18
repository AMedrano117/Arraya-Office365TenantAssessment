#Manages and processes user batches either in total or in specified segments.
function Manage-UserBatches {
    param (
        [Parameter(Mandatory)]
        [array]$UserArray,
        [Parameter(Mandatory=$false, HelpMessage="Include this switch to process all items in the array automatically. Omit this switch to enable specifying the number of batch jobs and the specific job number to process, allowing for segmented processing of the array.")]
        [switch]$ProcessAll
    )
    
    # Check if ProcessAll switch is used, if not, prompt for it
    if (!$PSBoundParameters.ContainsKey('ProcessAll')) {
        $ProcessAll = [bool]::Parse((Read-Host "Process all items? (true/false)"))
    }
    
    if ($ProcessAll) {
        Write-Host "Processing all users. Total Users: $($UserArray.count)"
        return $UserArray
    }
    else {
        $maxJobs = 0
        while ($maxJobs -le 0) {
            $inputMaxJobs = Read-Host "Enter max jobs (integer)"
            # Validate input for integer
            if (![int]::TryParse($inputMaxJobs, [ref]$maxJobs)) {
                Write-Host "Invalid input. Please enter a valid integer for max jobs."
                $maxJobs = 0
            }
        }

        $JobNumber = 0
        while ($JobNumber -le 0 -or $JobNumber -gt $maxJobs) {
            $inputJobNumber = Read-Host "Enter job number (integer)"
            # Validate input for integer
            if (![int]::TryParse($inputJobNumber, [ref]$JobNumber) -or $JobNumber -le 0 -or $JobNumber -gt $maxJobs) {
                Write-Host "Invalid job number. Please enter a job number between 1 and $maxJobs."
                $JobNumber = 0
            }
        }
        
        $recipientsPerJob = [math]::Ceiling($UserArray.count / $maxJobs)
        Write-Host "Users per job: $recipientsPerJob"
        
        $start = ($JobNumber - 1) * $recipientsPerJob
        $end = [math]::Min(($start + $recipientsPerJob - 1), ($UserArray.count - 1))
        
        $batchGroup = @($UserArray[$start..$end])
        Write-Host "Processing batch job number $JobNumber with users from index $start to $end."
        return $batchGroup
    }
}