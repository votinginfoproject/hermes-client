#---------------------------------------------
# Zip functions modified from http://www.technologytoolbox.com/blog/jjameson/archive/2012/02/28/zip-a-folder-using-powershell.aspx
#---------------------------------------------

If ((Test-Path variable:\PSVersionTable) -And ($PSVersionTable.PSVersion.Major -lt 4)) {
   Throw "You are running an old version of PowerShell. Please update to at least version 4." +
         "`r`n" + "Please see the following link:" +
         "`r`n" + "http://social.technet.microsoft.com/wiki/contents/articles/21016.how-to-install-windows-powershell-4-0.aspx"
}

function CountZipItems(
    [__ComObject] $zipFile)
{
    If ($zipFile -eq $null)
    {
        Throw "Value cannot be null: zipFile"
    }

    Write-Host ("Counting files in zip file (" + $zipFile.Self.Path + ")...")

    [int] $count = CountZipItemsRecursive($zipFile)

    Write-Host ($count.ToString() + " items in zip file (" `
        + $zipFile.Self.Path + ").")

    return $count
}

function CountZipItemsRecursive(
    [__ComObject] $parent)
{
    If ($parent -eq $null)
    {
        Throw "Value cannot be null: parent"
    }

    [int] $count = 0

    $parent.Items() |
        ForEach-Object {
            If ($_.IsFolder -eq $true) {
                $count += CountZipItemsRecursive($_.GetFolder)
            }
            Else {
                $count += 1
            }
        }

    return $count
}

function IsFileLocked(
    [string] $path)
{
    If ([string]::IsNullOrEmpty($path) -eq $true)
    {
        Throw "The path must be specified."
    }

    [bool] $fileExists = Test-Path $path

    If ($fileExists -eq $false)
    {
        Throw "File does not exist (" + $path + ")"
    }

    [bool] $isFileLocked = $true

    $file = $null

    Try
    {
        $file = [IO.File]::Open(
            $path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::None)

        $isFileLocked = $false
    }
    Catch [IO.IOException]
    {
        If ($_.Exception.Message.EndsWith(
            "it is being used by another process.") -eq $false)
        {
            Throw $_.Exception
        }
    }
    Finally
    {
        If ($file -ne $null)
        {
            $file.Close()
        }
    }

    return $isFileLocked
}

function GetWaitInterval(
    [int] $waitTime)
{
    If ($waitTime -lt 1000)
    {
        return 100
    }
    ElseIf ($waitTime -lt 5000)
    {
        return 1000
    }
    Else
    {
        return 5000
    }
}

function WaitForZipOperationToFinish(
    [__ComObject] $zipFile,
    [int] $expectedNumberOfItemsInZipFile)
{
    If ($zipFile -eq $null)
    {
        Throw "Value cannot be null: zipFile"
    }
    ElseIf ($expectedNumberOfItemsInZipFile -lt 1)
    {
        Throw "The expected number of items in the zip file must be specified."
    }

    Write-Host -NoNewLine "Waiting for zip operation to finish..."
    Start-Sleep -Milliseconds 100

    [int] $waitTime = 0
    [int] $maxWaitTime = 60 * 1000
    while($waitTime -lt $maxWaitTime)
    {
        [int] $waitInterval = GetWaitInterval($waitTime)

        Write-Host -NoNewLine "."
        Start-Sleep -Milliseconds $waitInterval
        $waitTime += $waitInterval

        Write-Debug ("Wait time: " + $waitTime / 1000 + " seconds")

        [bool] $isFileLocked = IsFileLocked($zipFile.Self.Path)

        If ($isFileLocked -eq $true)
        {
            Write-Debug "Zip file is locked by another process."
            Continue
        }
        Else
        {
            Break
        }
    }

    Write-Host

    If ($waitTime -ge $maxWaitTime)
    {
        Throw "Timeout exceeded waiting for zip operation"
    }

    [int] $count = CountZipItems($zipFile)

    If ($count -eq $expectedNumberOfItemsInZipFile)
    {
        Write-Debug "The zip operation completed succesfully."
    }
    ElseIf ($count -eq 0)
    {
        Throw ("Zip file is empty. This can occur if the operation is" `
            + " cancelled by the user.")
    }
    ElseIf ($count -gt $expectedCount)
    {
        Throw "Zip file contains more than the expected number of items."
    }
}

function ZipFolder(
    [IO.DirectoryInfo] $directory,
    [string] $zipFileName)
{
    If ($directory -eq $null)
    {
        Throw "Value cannot be null: directory"
    }

    Write-Host ("Creating zip file for folder (" + $directory.FullName + ")...")

    [IO.DirectoryInfo] $parentDir = $directory.Parent

    If (Test-Path $zipFileName)
    {
        Throw "Zip file already exists ($zipFileName)."
    }

    Set-Content $zipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))

    $shellApp = New-Object -ComObject Shell.Application
    $zipFile = $shellApp.NameSpace($zipFileName)

    If ($zipFile -eq $null)
    {
        Throw "Failed to get zip file object."
    }

    [int] $expectedCount = (Get-ChildItem $directory -File -Recurse).Count
    Write-Host ("Found $expectedCount files in " + $directory.FullName)

    If ($expectedCount -eq 0) {
        Throw ("Unable to build zip file; the folder " + $directory.FullName + " is empty")
    }

    $zipFile.CopyHere($directory.FullName)

    WaitForZipOperationToFinish $zipFile $expectedCount

    Write-Host -Fore Green ("Successfully created zip file for folder (" `
        + $directory.FullName + ").")
}

function FTP-Upload
{
    param(
        [string] $user = $null,
        [string] $pass = $null,
        [string] $file = $null,
        [string] $server = $null,
        [string] $zipFilename = $null
    )

    # Create FTP Rquest Object
    $FTPRequest = [System.Net.FtpWebRequest]::Create($server+$zipFilename)
    $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $FTPRequest.Credentials = new-object System.Net.NetworkCredential($user, $pass)
    $FTPRequest.EnableSsl = $true
    $FTPRequest.UseBinary = $true
    $FTPRequest.UsePassive = $true
    $FTPRequest.KeepAlive = $false

    # The file stream for upload
    $FileStream = new-object System.IO.FileStream($file, [System.IO.FileMode]'Open')

    # The FTP stream
    $FTPStream = $FTPRequest.GetRequestStream()

    $FileStream.CopyTo($FTPStream)

    # Cleanup
    $FileStream.Close()
    $FileStream.Dispose()
    $FTPStream.Close()
    $FTPStream.Dispose()
}

#---------------------------------------------
# Deletes file if exists
#---------------------------------------------
function delete-if-exists
{
    If (Test-Path $args[0]){
     Remove-Item $args[0];
    }
}
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$username = "admin"
$password = "admin"
$server = "ftp://ftp.votinginfoproject.org/"

$fips = "12345"
$electionDate = "2016-11-08"

$zipFilename = "vipFeed-$fips-$electionDate.zip"

echo "Clearing old feed zips if they exist"
delete-if-exists $scriptPath\$zipFilename

echo "Creating new feed zip"
[IO.DirectoryInfo] $directory = Get-Item "$scriptPath\data"
ZipFolder $directory $scriptPath\$zipFilename

echo "Uploading feed file"
FTP-Upload -user $username -pass $password -file $scriptPath\$zipFilename -zipFilename $zipFilename -server $server
