<#
.DESCRIPTION
This script attempts to authenticate to OneDrive as client https://photos.onedrive.com 
(which has permissions to download LivePhotos) and downloads all LivePhotos at a given
location within your personal OneDrive.
.PARAMETER SaveTo
Target path where to save live photos.
.PARAMETER PathToScan
DOS-Style path on your OneDrive that should be scanned. Most likely '\Pictures\Camera Roll' or any other shared Camera Roll folder.

.EXAMPLE
.\Download-ODLivePhotos.ps1 'C:\Live Photos'

.NOTES
Author: Petr Vyskocil

There is no error checking, so it is recommended to re-run the command on bigger libraries.
If there are any errors during the download (OneDrive sometimes fails randomly with error "Our
services aren't available right now. We're working to restore all services as soon as possible.
Please check back soon."), next run will download the missing files and skip already downloaded
ones.
#>
param (
    [Parameter(Mandatory)]
    [string] $SaveTo,
    [string] $AccessToken,
    [string] $PathToScan = '\Pictures\Camera Roll'
)


function Register-WebView2Type
{
    <#
    .DESCRIPTION
    Check that WebView2 exists, if not obtain it from nuget.

    .EXAMPLE
    Register-WebView2Type
    .NOTES
    Author: Petr Vyskocil
    #>
    $Package = "Microsoft.Web.WebView2"
    $Version = "1.0.3124.44"
    $BasePath = "$env:temp\ODLivePhotos\"

    if (!("Microsoft.Web.WebView2.WinForms.WebView2" -as [type])) {
        if (!(Test-Path "$($BasePath)\Microsoft.Web.WebView2.WinForms.dll")) {
            Write-Output "  Downloading nuget package $($Package) $($Version)"
            Install-Package -Source "https://www.nuget.org/api/v2" -Name $Package -RequiredVersion $Version -Scope CurrentUser -Destination $BasePath -Force
            Write-Output "  Copying package files to script directory"
            foreach ( $Framework in (Get-ChildItem "$BasePath\$($Package).$($Version)\lib" -Directory) ) { 
                Copy-Item -Recurse -Path "$BasePath\$($Package).$($Version)\lib\$($Framework)\*.dll" -Destination $BasePath -Force
            }
            $Arch = (Get-CimInstance Win32_operatingsystem).OSArchitecture
            if ($Arch -eq "32-bit") {
                Copy-Item -Recurse -Path  "$BasePath\$($Package).$($Version)\runtimes\win-x86\native\*.dll" -Destination $BasePath -Force
            }
            if ($Arch -eq "64-bit") {
                Copy-Item -Recurse -Path  "$BasePath\$($Package).$($Version)\runtimes\win-x64\native\*.dll" -Destination $BasePath -Force
            }
            if ($Arch -eq "ARM 64-bit Processor") {
                Copy-Item -Recurse -Path  "$BasePath\$($Package).$($Version)\runtimes\win-arm64\native\*.dll" -Destination $BasePath -Force
            }
            Remove-Item -Recurse -Force "$BasePath\$($Package).$($Version)"
        }
        Write-Output "  Registering WebView2 type"
        Add-Type -Path "$BasePath\Microsoft.Web.WebView2.WinForms.dll"
    }
}

function Get-ODPhotosToken
{
    <#
    .DESCRIPTION
    Connect to OneDrive for authentication with a OneDrive web Photos client.

    Open OneDrive photos in WebView2 and steal authentication token once
    authenticated. This is a new SharePoint-like authentication that can
    be requested only by MSFT internal apps, unfortunately the public
    MS Graph API does not have access to Live Photos.

    .EXAMPLE
    $access_token=Get-ODPhotosToken
    Connect to OneDrive for authentication and save the token to $access_token
    .NOTES
    Author: Petr Vyskocil
    #>
    $Hash = [hashtable]::Synchronized(@{}) 
	$Env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--enable-features=msSingleSignOnOSForPrimaryAccountIsShared"
    $Web = New-Object Microsoft.Web.WebView2.WinForms.WebView2
    $Web.CreationProperties = New-Object Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
    $Web.CreationProperties.UserDataFolder = "$env:temp\ODLivePhotos\"
    $Web.Dock = "Fill"
    $Web.source  = "https://onedrive.live.com/?qt=allmyphotos&photosData=%2F&sw=bypassConfig&v=photos"
    $Web.add_CoreWebView2InitializationCompleted({
        $Web.CoreWebView2.Settings.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36 Edg/134.0.3124.85'
        $Web.CoreWebView2.add_WebResourceResponseReceived({
            param($WebView2, $E)
            if ($E.Request.Uri.StartsWith('https://my.microsoftpersonalcontent.com/_api/') -or $E.Request.Uri.StartsWith('https://api.onedrive.com/')) {
                Write-Host $E.Request.Uri
                if ($E.Request.Headers.Contains('Authorization')) {
                    $Hash.AuthToken = $E.Request.Headers.GetHeader('Authorization')
                    $Form.Close()
                }
            }
        })
    })
    $Form = New-Object System.Windows.Forms.Form -Property @{Width=800;Height=800;Text="OneDrive Live Photo Downloader - Authentication Token capture dialog"} -ErrorAction Stop
    $Form.Controls.Add($web)
    $Form.Add_Shown( { $form.Activate() } )
    $Form.ShowDialog() | Out-Null

    $Web.Dispose()

    return $Hash.AuthToken    
}

function Download-LivePhotosAuth
{
    <#
    .DESCRIPTION
    Download all Live Photos from a given OneDrive path
    .PARAMETER AccessToken
    Access token for OneDrive API that has ability to download live photos.
    .PARAMETER SaveTo
    Target path where to save live photos.
    .PARAMETER PathToScan
    DOS-Style path on your OneDrive that should be scanned. Most likely '\Pictures\Camera Roll' or any other shared Camera Roll filder you see.
    .PARAMETER CurrentPath
    Internal, used for recursion
    .PARAMETER ElementId
    Start at folder id (probably only useful internally for recursion too)
    .PARAMETER Uri
    Start processing on this URI (again for recursion)

    .EXAMPLE
    Download-LivePhotosAuth -AccessToken $token -SaveTo 'C:\LivePhotos' -PathToScan '\Pictures\Camera Roll'
    .NOTES
    Author: Petr Vyskocil
    #>
    PARAM(
        [Parameter(Mandatory=$True)]
        [string]$AccessToken,
        [Parameter(Mandatory=$True)]
        [string]$SaveTo,
        [string]$PathToScan='\',
        [string]$CurrentPath='',
        [string]$ElementId='',
        [string]$Uri=''
    )
    if (!$PathToScan.EndsWith('\')) { $PathToScan = $PathToScan + '\' }
    if (!$PathToScan.StartsWith('\')) { $PathToScan = '\' + $PathToScan }
    if ($Uri -eq '') {
        if ($ElementId -eq '') { 
            $CurrentPath='\'
            $Location='root'
        } else {
            $Location='items/' + $ElementId
        }
        $Uri = 'https://api.onedrive.com/v1.0/drive/' + $Location + '/children?%24filter=photo%2FlivePhoto+ne+null+or+folder+ne+null+or+remoteItem+ne+null&select=fileSystemInfo%2Cphoto%2Cid%2Cname%2Csize%2Cfolder%2CremoteItem'
    }
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $WebRequest=Invoke-WebRequest -Method 'GET' -Header @{ Authorization = $AccessToken} -ErrorAction SilentlyContinue -Uri $Uri
    $Response = ConvertFrom-Json $WebRequest.Content
    $Response.value | % {
        $FolderPath = $CurrentPath + $_.name + '\'
        if ([bool]$_.PSObject.Properties['folder']) {
            if ($FolderPath.StartsWith($PathToScan) -or $PathToScan.StartsWith($FolderPath)) { # We're traversing the target folder or we're getting into it
                Write-Output("Checking folder $($_.id) - $($FolderPath)")
                Download-LivePhotosAuth -AccessToken $AccessToken -SaveTo $SaveTo -PathToScan $PathToScan -CurrentPath $FolderPath -ElementId $_.id
            }
        }
        if ([bool]$_.PSObject.Properties['remoteItem']) {
            if ($FolderPath.StartsWith($PathToScan) -or $PathToScan.StartsWith($FolderPath)) { # We're traversing the target folder or we're getting into it
                Write-Output("Checking shared folder $($_.remoteItem.id) - $($FolderPath)")
                Download-LivePhotosAuth -AccessToken $AccessToken -SaveTo $SaveTo -PathToScan $PathToScan -CurrentPath $FolderPath -ElementId $_.remoteItem.id
            }
        }
        if ([bool]$_.PSObject.Properties['photo']) {
            if ([bool]$_.photo.PSObject.Properties['livePhoto']) {
                if ($CurrentPath.StartsWith($PathToScan)) {
                    $TargetPath = $SaveTo + '\' + $CurrentPath.Substring($PathToScan.Length)
                    if ( (Test-Path($TargetPath+$_.name)) -and # Target image exists
                         (Test-Path($TargetPath+([io.fileinfo]$_.name).basename+'.mov')) -and # Target video exists
                         (((Get-Item($TargetPath+$_.name)).Length + (Get-Item($TargetPath+([io.fileinfo]$_.name).basename+'.mov')).Length) -eq $_.size) # size of image and video together is onedrive's size
                       ) {
                        Write-Output "Live photo $($_.id) - $($CurrentPath + $_.name) already exists at $($TargetPath) - skipping."
                    } else {
                        Write-Output("Detected live photo $($_.id) - $($CurrentPath + $_.name). Saving image/video pair to $($TargetPath)")
                        Download-SingleLivePhoto -AccessToken $AccessToken -ElementId $_.id -SaveTo $TargetPath -ExpectedSize $_.size -LastModified $_.fileSystemInfo.lastModifiedDateTime
                    }
                }
            }
        }
    }
    if ([bool]$Response.PSobject.Properties["@odata.nextLink"]) 
    {
        Write-Debug("Getting more elements form service (@odata.nextLink is present)")
        Download-LivePhotosAuth -AccessToken $AccessToken -SaveTo $SaveTo -PathToScan $PathToScan -CurrentPath $CurrentPath -Uri $Response.'@odata.nextLink'
    }
}

function Download-SingleLivePhoto
{
    <#
    .DESCRIPTION
    Download single LivePhoto given it's ElementId and static data.
    .PARAMETER AccessToken
    Access token for OneDrive API that has ability to download live photos.
    .PARAMETER ElementId
    OneDrive ElementId of a LivePhoto
    .PARAMETER SaveTo
    Target path where to save live photos.
    .PARAMETER ExpectedSize
    Sum of photo and video file sizes, as reported in the containing folder
    .PARAMETER LastModified
    Date to set on a created file.

    .NOTES
    Author: Petr Vyskocil
    #>
    PARAM(
        [Parameter(Mandatory=$True)]
        [string]$AccessToken,
        [Parameter(Mandatory=$True)]
        [string]$ElementId,
        [Parameter(Mandatory=$True)]
        [string]$SaveTo,
        [Parameter(Mandatory=$True)]
        [int]$ExpectedSize,
        [Parameter(Mandatory=$True)]
        [datetime]$LastModified
    )
    
    if (!(Test-Path $SaveTo)) { New-Item -ItemType Directory -Force $SaveTo | Out-Null }
    
    # video part
    $Uri = "https://api.onedrive.com/v1.0/drive/items/$($ElementId)/content?format=video"
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $TmpFile = $SaveTo+'tmp-file.mov'
    $WebRequest=Invoke-WebRequest -Method "GET" -Uri $Uri -Header @{ Authorization = $AccessToken } -ErrorAction SilentlyContinue -OutFile $TmpFile -PassThru
    $ActualSize = $WebRequest.RawContentLength
    $FileName = ($WebRequest.Headers.'Content-Disposition'.Split('=',2)[-1]).Trim('"')
    if ($FileName) {
        Write-Debug("Renaming $TmpFile to $FileName")
        if (Test-Path($SaveTo+$FileName)) { Remove-Item ($SaveTo+$FileName) }
        Rename-Item -Path $TmpFile -NewName $FileName
        (Get-Item ($SaveTo+$FileName)).LastWriteTime = $LastModified
    }
    

    # image part
    $Uri = "https://api.onedrive.com/v1.0/drive/items/$($ElementId)/content"
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $TmpFile = $SaveTo+'tmp-file.img'
    $WebRequest=Invoke-WebRequest -Method "GET" -Uri $Uri -Header @{ Authorization = $AccessToken } -ErrorAction SilentlyContinue -OutFile $TmpFile -PassThru
    $ActualSize = $ActualSize + $WebRequest.RawContentLength
    $FileName = ($WebRequest.Headers.'Content-Disposition'.Split('=',2)[-1]).Trim('"')
    if ($FileName) {
        Write-Debug("Renaming $TmpFile to $FileName")
        if (Test-Path($SaveTo+$FileName)) { Remove-Item ($SaveTo+$FileName) }
        Rename-Item -Path $TmpFile -NewName $FileName
        (Get-Item ($SaveTo+$FileName)).LastWriteTime = $LastModified
    }
    
    if ($ActualSize -ne $ExpectedSize) { Write-Error("Error saving live photo $ElementId. Got $ActualSize bytes, expected $ExpectedSize bytes.") }

}


Write-Output "Live Photo downloader - Downloads Live Photos from OneDrive camera roll as saved by OneDrive iOS app."
Write-Output "(C) 2024 Petr Vyskocil. Licensed under MIT license."
Write-Output ""


# This disables powershell progress indicators, speeding up Invoke-WebRequest with big results by a factor of 10 or so
$ProgressPreference = 'SilentlyContinue'

if ($AccessToken -eq '') {
    Write-Output "Creating WebView2 component..."
    Register-WebView2Type
    Write-Output "Getting OneDrive Authentication token..."
    $AccessToken = Get-ODPhotosToken
}

Write-Output "Downloading Live Photos..."
Download-LivePhotosAuth -AccessToken $AccessToken -PathToScan $PathToScan -SaveTo $SaveTo
