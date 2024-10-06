# Set the path for the CSV file
$csvPath = '/Users/gabrielsilvestre/Downloads/gabriel.csv - Página1.csv'

# Load the CSV file into a variable
$users = Import-Csv -Path $csvPath

# Create an array to store processed user data
$processedUsers = @()

# Process each user from the CSV
foreach ($user in $users) {
    # Get the first and last names from displayName
    $displayName = $user.displayName
    $FirstName = ""
    $LastName = ""

    if ($displayName -match ',') {
        # If the name has a comma, it's in "LastName, FirstName" format
        $nameParts = $displayName -split ','
        $LastName = $nameParts[0].Trim()
        $FirstName = $nameParts[1].Trim()
    } else {
        # Otherwise, it's probably "FirstName LastName"
        $nameParts = $displayName -split ' '
        $FirstName = $nameParts[0]
        $LastName = $nameParts[-1]
    }

    # Fix the "lastlogontimestamp" issue (it uses commas instead of periods, so let's convert it)
    $timestampString = $user.lastlogontimestamp -replace ',', '.'
    $LastLogonDate = [DateTime]::FromFileTimeUtc([Int64][double]::Parse($timestampString, [System.Globalization.CultureInfo]::InvariantCulture))

    # Parse and format the hire date
    $HireDate = [DateTime]::Parse($user.HireDate)
    $DateFormat = "M/d/yyyy h:mm:ss tt"
    $formattedLogonDate = $LastLogonDate.ToString($DateFormat)
    $formattedHireDate = $HireDate.ToString($DateFormat)

    # Calculate how many days since the last logon
    $daysSinceLogon = (Get-Date) - $LastLogonDate
    $days = [int]$daysSinceLogon.TotalDays

    # Figure out the logon interval based on how many days it's been since the last logon
    switch ($days) {
        { $_ -lt 30 } { $logonInterval = "Less than 30 days"; break }
        { $_ -lt 60 } { $logonInterval = "Between 30 and 60 days"; break }
        { $_ -lt 90 } { $logonInterval = "Between 60 and 90 days"; break }
        { $_ -lt 180 } { $logonInterval = "Between 90 and 180 days"; break }
        { $_ -lt 365 } { $logonInterval = "Between 180 and 365 days"; break }
        default { $logonInterval = "More than 1 year" }
    }

    # Create an object with all the information for this user
    $userInfo = [PSCustomObject]@{
        FirstName = $FirstName
        LastName = $LastName
        displayName = $user.displayName
        mail = $user.mail
        HireDate = $formattedHireDate
        LastLogonDate = $formattedLogonDate
        LastLogonInterval = $logonInterval
    }

    # Add this user's info to the processed users array
    $processedUsers += $userInfo
}

# Define output paths for CSV and HTML files
$outputDir = Split-Path -Path $csvPath -Parent
$outputCsvPath = Join-Path -Path $outputDir -ChildPath "output.csv"
$outputHtmlPath = Join-Path -Path $outputDir -ChildPath "output.html"

# Export the processed data to CSV and HTML files
$processedUsers | Export-Csv -Path $outputCsvPath -NoTypeInformation
$processedUsers | ConvertTo-Html -Property FirstName, LastName, displayName, mail, HireDate, LastLogonDate, LastLogonInterval -Title "User Information" | Set-Content -Path $outputHtmlPath

# Let the user know we're done
Write-Host "Processing complete! Output files are in the same directory as the input CSV:"
Write-Host "CSV File: $outputCsvPath"
Write-Host "HTML File: $outputHtmlPath"