param([string]$Arguments)

$startDirectory         = $Arguments + '\'

$numberOfTopDirectories = 5
$numberOfTopFiles       = 5
$emailTo                = 's.youssef@wic-sa.com'
$smtpSrv                = 'notif-scom@wic-sa.com'
$emailFrom              = 'diskSpaceDetail@wic-sa.com'

#region Get-Metadata

$timeZone               =  ([TimeZoneInfo]::Local).Id
$scanDate               = Get-Date -Format 'yyyy-MM-dd hh:MM:ss'
$WindowsVersion         = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption

try {
    $computerDescription  = Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\LanmanServer\Parameters | Select-Object -ExpandProperty srvcomment
} catch {
    $computerDescription  = ''
}

try {
    $adSearcher           = New-Object System.DirectoryServices.DirectorySearcher
    $adSearcher.Filter    = "(&(objectCategory=computer)(cn=$env:computername))"
    $adComputer           = $adSearcher.FindOne()
    $adComputerProperties = $adComputer | Select-Object -ExpandProperty Properties
    $adComputerLdapPath   = $adComputerProperties.distinguishedname
} catch {
    $adComputerLdapPath   = 'Failed to extract AD Information.' + $_.StackTrace
}


$diskDetails    = Get-WMIObject -Namespace root/cimv2 -Class Win32_LogicalDisk | Where-Object {$_.DeviceID -match "$($Arguments)" } | Select-Object -Property Size, FreeSpace, VolumeName, DeviceID
$DriveLetter    = $diskDetails.DeviceID
$DriveName      = $diskDetails.VolumeName
$SizeInGB       = "{0:0}" -f ($diskDetails.Size/1GB)
$FreeSpaceInGB  = "{0:0}" -f ($diskDetails.FreeSpace/1GB)
$PercentFree    = "{0:0}" -f (($diskDetails.FreeSpace / $diskDetails.Size) * 100)

$diskMessage    = "<table><tr><td>$($DriveLetter)</td><td> ($($DriveName))</td></tr>"
$diskMessage   += "<tr><td>Total: $($SizeInGB) GB |</td><td>Free: $($FreeSpaceInGB) GB ($($PercentFree) %)</td></tr></table>"

#endregion Get-Metadata


Function Get-BigDirectories {

    param(
        [string]$startDirectory,
        [ref]$sortedDirectories
    )

    $bigDirList = New-Object -TypeName System.Collections.ArrayList

    if (Test-Path -Path $startDirectory) {

        & "$env:ComSpec" /c dir $startDirectory /-c /s | ForEach-Object {

                $null      = $_ -match 'Directory\s{1}of\s{1}(?<dirName>[\w:\\\s\.\-\(\)_#{}\$\%\+\[\]]{1,})'
                $dirName   = $Matches.dirName

                $null      = $_ -match '\s{1,}\d{1,}\sFile\(s\)\s{1,}(?<lengh>\d{1,})'
                $dirLength = $Matches.lengh

                if ($dirName -and $dirLength) {

                    $dirLength = [float]::Parse($dirLength)
                    $myFileHsh = @{'Name'=([string]$dirName)}
                    $myFileHsh.Add('Length',$dirLength)
                    $myFileObj = New-Object -TypeName PSObject -Property $myFileHsh
                    $null = $bigDirList.Add($myFileObj)

                }

                $dirName = ''
                $dirLength = 0

        } #END cmd /c dir C:\windows /-c /s | ForEach-Object

    } else {

        if ($startDirectory) {
            $dirName   = 'Error'
            $dirLength = 'No directory passed in.'
        } else  {
            $dirName   = $startDirectory
            $dirLength = $error.Message.ToString()
        }

        $dirLength = [float]::Parse($dirLength)
        $myFileHsh = @{'Name'=([string]$dirName)}
        $myFileHsh.Add('Length',$dirLength)
        $myFileObj = New-Object -TypeName PSObject -Property $myFileHsh

        $null = $bigDirList.Add($myFileObj)

    } #END if (Test-Path -Path $startDirectory)

    $sortedDirectories.Value = $bigDirList

} #End Function Get-BigDirectories


Function Convert-LengthToReadable {

    param(
        [System.Collections.ArrayList]$lengthList,
        [ref]$readableList
    )

    $allFiles = New-Object -TypeName System.Collections.ArrayList

    $lengthList | ForEach-Object {

        $sizeRaw  = $_.Length
        $sizeUnit = 'KB'
        $fileSize = 0

        if ($sizeRaw -gt 1kb -and $sizeRaw -lt 1mb ) {
            $fileSize = $sizeRaw / 1mb
            $sizeUnit = 'MB'
        } elseif ($sizeRaw -gt 1mb -and $sizeRaw -lt 1gb ) {
            $fileSize = $sizeRaw / 1mb
            $sizeUnit = 'MB'
        } elseif ($sizeRaw -gt 1gb -and $sizeRaw -lt 1tb ) {
            $fileSize = $sizeRaw / 1gb
            $sizeUnit = 'GB'
        } elseif ($sizeRaw -gt 1tb ) {
            $fileSize = $sizeRaw / 1tb
            $sizeUnit = 'TB'
        } else {
            $fileSize = $sizeRaw
            $sizeUnit = 'KB?'
        }

        $fileSize = [Math]::Round($fileSize,2)
        $myFileHsh = @{'Name'=([string]$_.Name)}
        $myFileHsh.Add('fileSize',([float]$fileSize))
        $myFileHsh.Add('sizeUnit',([string]$sizeUnit))
        $myFileHsh.Add('Length',([float]$sizeRaw))
        $myFileHsh.Add('LastWriteTime',$_.LastWriteTime)
        $myFileObj = New-Object -TypeName PSObject -Property $myFileHsh
        $null = $allFiles.Add($myFileObj)

    }

    $readableList.Value = $allFiles

} #End Function Convert-LengthToReadable


Function Send-TopDirectoryMailReport {

    param(
        [string]$diskMetaData,
        [string]$runInfo,
        [System.Array]$tDirectories,
        [System.Collections.Hashtable]$tDirsAndFiles,
        [System.Collections.Hashtable]$nDirsAndFiles
    )

    $directoryDetails   = ''
    $dirAndFilesDetails = ''
    $dirAndNewFilesDetails = ''

    foreach ($dirItem in $tDirectories) {

        $dirAndFilesDetails    += "<tr><td style=`"font-weight: bold; text-align:center`"><br />&nbsp;$($dirItem.Name)</td></tr>"
        $dirAndNewFilesDetails += "<tr><td style=`"font-weight: bold; text-align:center`"><br />&nbsp;$($dirItem.Name)</td></tr>"
        $directoryDetails      += "<tr><td>$($dirItem.Name)</td><td>$($dirItem.fileSize)</td><td>$($dirItem.sizeUnit)</td><tr>"

        $matchEntry = $tDirsAndFiles.($dirItem.Name)
        $matchEntry | ForEach-Object {
            $dirAndFilesDetails += "<tr><td>$($_.Name)</td><td>$($_.fileSize)</td><td>$($_.sizeUnit)</td><td>$($_.LastWriteTime)</td><tr>"
        }

        $matchEntry = $nDirsAndFiles.($dirItem.Name)
        $matchEntry | ForEach-Object {
            $dirAndNewFilesDetails += "<tr><td>$($_.Name)</td><td>$($_.fileSize)</td><td>$($_.sizeUnit)</td><td>$($_.LastWriteTime)</td><tr>"
        }

    } #End     foreach ($dirItem in $tDirectories)

    $htmlBegin   = "<!DOCTYPE html><html><head><title>DISK FULL - Troubleshooting Assistance on $($env:Computername) -  $($computerDescription)</title>"
    $htmlBegin  += "<h1><span style=`"background-color:#D3D3D3`">DISK FULL - Troubleshooting Assistance on $($env:Computername)</span></h1></head>"
    $htmlBegin  += "<h3><span style=`"background-color:#D3D3D3`"> $($computerDescription) </span></h3></head>"

    $htmlMiddle  = '<body style="color:#000000; font-size:12pt;"><p>&nbsp;</p><span style="color:#B22222; font-weight: bold; background-color:#D3D3D3; font-size:14pt;">Disk details:</span><br />' + $diskMetaData + '<p>&nbsp;</p>'
    $htmlMiddle += '<span style="color:#000080; font-weight: bold; background-color:#D3D3D3; font-size:14pt;">Largest Directories:</span><br /><br />&nbsp;<table>' + $directoryDetails + '</table><p>&nbsp;</p>'
    $htmlMiddle += '<span style="color:#000080; font-weight: bold; background-color:#D3D3D3; font-size:14pt;">Newest files:</span><br /><table>' + $dirAndNewFilesDetails + '</table><p>&nbsp;</p>'
    $htmlMiddle += '<span style="color:#000080; font-weight: bold; background-color:#D3D3D3; font-size:14pt;">Largest files:</span><br /><table>' + $dirAndFilesDetails + '</table><p>&nbsp;</p>'
    $htmlMiddle += '<span style="color:#FF8C00; font-weight: bold; background-color:#D3D3D3; font-size:14pt;">Meta Information:<br /></span>' + $runInfo
    $htmlEnd     = '</body></html>'

    $htmlCode = $htmlBegin + $htmlMiddle + $htmlEnd

    $mailMessageParms = @{
        To          = $emailTo
        From        = $emailFrom
        Subject     = "DISK Full - Troubleshooting Assistance on $($env:computername) . $($computerDescription)"
        Body        = $htmlCode
        Smtpserver  = $smtpSrv
        ErrorAction = "SilentlyContinue"
        BodyAsHTML  = $true
    }

    Send-MailMessage @mailMessageParms

} #End Send-TopDirectoryMailReport


$startTime = Get-Date

$bigDirList = New-Object -TypeName System.Collections.ArrayList
Get-BigDirectories -startDirectory $startDirectory -sortedDirectories ([ref]$bigDirList)

$bigDirListReadable = New-Object -TypeName System.Collections.ArrayList
Convert-LengthToReadable -lengthList $bigDirList -readableList ([ref]$bigDirListReadable)

$topDirectories  = $bigDirListReadable | Sort-Object -Property Length -Descending | Select-Object -First $numberOfTopDirectories
$topDirsAndFiles = New-Object -TypeName System.Collections.Hashtable

$topDirsAndNewestFiles = New-Object -TypeName System.Collections.Hashtable

foreach ($tDirectory in $topDirectories) {

    $tDirName       = $tDirectory.Name

    $tmpFileList    = New-Object -TypeName System.Collections.ArrayList
    $tmpNewFileList = New-Object -TypeName System.Collections.ArrayList
    $newFilesInDir  = New-Object -TypeName System.Collections.ArrayList

    $filesInTDir    = Get-ChildItem -Path $tDirName | Where-Object { $_.PSIsContainer -eq $false } | Select-Object -Property DirectoryName, Name, LastWriteTime, Length
    $filesInTDir      | ForEach-Object {
        $null = $newFilesInDir.Add($_)
    }

    $filesInTDir    = $filesInTDir   | Sort-Object -Property Length        -Descending | Select-Object -First $numberOfTopFiles
    $newFilesInDir  = $newFilesInDir | Sort-Object -Property LastWriteTime -Descending | Select-Object -First $numberOfTopFiles

    $filesInTDir | Select-Object -Property DirectoryName, Name, LastWriteTime, Length | ForEach-Object {
        $null = $tmpFileList.Add($_)
    }

    $newFilesInDir | Select-Object -Property DirectoryName, Name, LastWriteTime, Length | ForEach-Object {
        $null = $tmpNewFileList.Add($_)
    }

    $bigFileList = New-Object -TypeName System.Collections.ArrayList
    Convert-LengthToReadable -lengthList $tmpFileList -readableList ([ref]$bigFileList)

    $newFileList = New-Object -TypeName System.Collections.ArrayList
    Convert-LengthToReadable -lengthList $tmpNewFileList -readableList ([ref]$newFileList)

    $topDirsAndFiles.Add($tDirName,$bigFileList)
    $topDirsAndNewestFiles.Add($tDirName,$newFileList)

} #End foreach ($tDirectory in $topDirectories)

$endTime      = Get-Date
$requiredTime = New-TimeSpan -Start $startTime -End $endTime
$reqHours     = [Math]::Round($requiredTime.TotalHours,2)
$reqMinutes   = [Math]::Round($requiredTime.TotalMinutes,2)
$reqSeconds   = [Math]::Round($requiredTime.TotalSeconds,2)

$metaInfo     = "<table><tr><td>Gathering details took:</td><td>$($reqHours) Hours / $($reqMinutes) Minutes / $($reqSeconds) Seconds.</td></tr>"
$metaInfo    += "<tr><td>Checking time:</td><td>$($scanDate), $($timeZone)</td></tr>"
$metaInfo    += "<tr><td>LDAP Path:</td><td>$($adComputerLdapPath)</td></tr>"
$metaInfo    += "<tr><td>Operating System:</td><td>$($WindowsVersion)</td></tr></table>"

$sendTopDirectoryMailReport = @{
    tDirectories  = $topDirectories
    tDirsAndFiles = $topDirsAndFiles
    nDirsAndFiles = $topDirsAndNewestFiles
    diskMetaData  = $diskMessage
    runInfo       = $metaInfo
}

Send-TopDirectoryMailReport @sendTopDirectoryMailReport