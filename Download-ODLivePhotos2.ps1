<#
.DESCRIPTION
This script attempts to authenticate to OneDrive as client https://photos.onedrive.com 
(which has permissions to download LivePhotos) and downloads all LivePhotos at a given
location within your personal OneDrive.
.PARAMETER SaveTo
Target path where to save live photos.
.PARAMETER AccessToken
OneDrive application access token. Currently required to be copied from browser.
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


function Get-ODPhotosToken
{
    <#
    .DESCRIPTION
	Does not work for migrated onedrive accounts anymore, please copy auth token from browser for the time being.
	
    Connect to OneDrive for authentication with a OneDrive web Photos client id. Adapted from https://github.com/MarcelMeurer/PowerShellGallery-OneDrive to mimic OneDrive Photos web OIDC login.
    Unfortunately using custom ClientId seems impossible - generic OD client IDs are missing the ability to download Live Photos.
    .PARAMETER ClientId
    ClientId of OneDrive Photos web app (073204aa-c1e0-4e66-a200-e5815a0aa93d)
    .PARAMETER Scope
    Comma-separated string defining the authentication scope (https://dev.onedrive.com/auth/msa_oauth.htm). Default: "OneDrive.ReadWrite,offline_access,openid,profile".
    .PARAMETER RedirectURI
    Code authentication requires a correct URI. Must be https://photos.onedrive.com/auth/login.

    .EXAMPLE
    $access_token=Get-ODPhotosToken
    Connect to OneDrive for authentication and save the token to $access_token
    .NOTES
    Author: Petr Vyskocil
    #>
    PARAM(
        [string]$ClientId = "073204aa-c1e0-4e66-a200-e5815a0aa93d",
        [string]$Scope = "OneDrive.ReadWrite,offline_access,openid,profile",
        [string]$RedirectURI ="https://photos.onedrive.com/auth/login",
        [switch]$DontShowLoginScreen=$false,
        [switch]$LogOut
    )
    $Authentication=""
    
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null
    [Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null
    if ($Logout)
    {
        $URIGetAccessToken="https://login.live.com/logout.srf"
    }
    else
    {
        $URIGetAccessToken="https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?client_id="+$ClientId+"&nonce=uv."+(New-Guid).Guid+"&response_mode=form_post&scope="+$Scope+"&response_type=code&redirect_URI="+$RedirectURI
    }
    $form = New-Object Windows.Forms.Form
    if ($DontShowLoginScreen)
    {
        write-debug("Logon screen suppressed by flag -DontShowLoginScreen")
        $form.Opacity = 0.0;
    }
    $form.text = "Authenticate to OneDrive"
    $form.size = New-Object Drawing.size @(700,600)
    $form.Width = 660
    $form.Height = 775
    $web=New-object System.Windows.Forms.WebBrowser
    $web.IsWebBrowserContextMenuEnabled = $true
    $web.Width = 600
    $web.Height = 700
    $web.Location = "25, 25"
    $web.ScriptErrorsSuppressed = $true
    $DocComplete  = {
        if ($web.Url.AbsoluteUri -match "access_token=|error|code=|logout|/auth/login") {$form.Close() } # 
    }
    $web.Add_DocumentCompleted($DocComplete)
    $form.Controls.Add($web)
    $web.navigate($URIGetAccessToken)
    $form.showdialog() | out-null
    $Authentication = New-Object PSObject
    # The returned code=XXXX is irrelevant, the actual secrets are sent as cookies:
    $web.Document.Cookie -split ';' | % { 
        $cookie = $_ -split '='
        $cookieValue = [uri]::UnescapeDataString($cookie[1])
        $Authentication | Add-Member NoteProperty $cookie[0].Trim() $cookieValue
    }
    
    if (-Not $Authentication.'AccessToken-OneDrive.ReadWrite') {
        write-error("Cannot get authentication token. This program does not suport token refresh at the moment. Try again, if you fail a few times, restart and try again.")
        # TODO: Refresh token should be handled here : GET https://photos.onedrive.com/auth/refresh?scope=OneDrive.ReadWrite&refresh_token=
        # Unfortunately it seems that this page hangs up the WebBrowser control, and we need all the cookies transferred to use...
        # Deemed not worth debugging for utility that runs for a short period of time
        return $web
        
    }
    
    return $Authentication
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
		[string]$DriveId='',
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
		if ($DriveId -eq '') {
			$Uri = 'https://my.microsoftpersonalcontent.com/_api/v2.1/drive/' + $Location + '/children?%24filter=photo%2FlivePhoto+ne+null+or+folder+ne+null+or+remoteItem+ne+null&select=fileSystemInfo%2Cphoto%2Cid%2Cname%2Csize%2Cfolder%2CremoteItem'
		} else {
			$Uri = 'https://my.microsoftpersonalcontent.com/_api/v2.1/drives/' +$DriveId + '/' + $Location + '/children?%24filter=photo%2FlivePhoto+ne+null+or+folder+ne+null+or+remoteItem+ne+null&select=fileSystemInfo%2Cphoto%2Cid%2Cname%2Csize%2Cfolder%2CremoteItem'
		}
    }
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $WebRequest=Invoke-WebRequest -Method 'GET' -Header @{ Authorization = "BEARER "+$AccessToken; Prefer = "Include-Feature=AddToOneDrive"} -ErrorAction SilentlyContinue -Uri $Uri
    $Response = ConvertFrom-Json $WebRequest.Content
    $Response.value | % {
        $FolderPath = $CurrentPath + $_.name + '\'
        if ([bool]$_.PSObject.Properties['folder']) {
            if ($FolderPath.StartsWith($PathToScan) -or $PathToScan.StartsWith($FolderPath)) { # We're traversing the target folder or we're getting into it
                Write-Output("Checking folder $($_.id) - $($FolderPath)")
                Download-LivePhotosAuth -AccessToken $AccessToken -SaveTo $SaveTo -PathToScan $PathToScan -CurrentPath $FolderPath -ElementId $_.id -DriveId $DriveId
            }
        }
        if ([bool]$_.PSObject.Properties['remoteItem']) {
            if ($FolderPath.StartsWith($PathToScan) -or $PathToScan.StartsWith($FolderPath)) { # We're traversing the target folder or we're getting into it
                Write-Output("Checking shared folder $($_.remoteItem.id) - $($FolderPath)")
                Download-LivePhotosAuth -AccessToken $AccessToken -SaveTo $SaveTo -PathToScan $PathToScan -CurrentPath $FolderPath -ElementId $_.remoteItem.id -DriveId $_.remoteItem.parentReference.driveId
            }
        }
        if ([bool]$_.PSObject.Properties['photo']) {
            if ([bool]$_.photo.PSObject.Properties['livePhoto']) {
                if ($CurrentPath.StartsWith($PathToScan)) {
                    $TargetPath = $SaveTo + '\' + $CurrentPath.Substring($PathToScan.Length)
					$VideoName = ([io.fileinfo]$_.name).basename+'.mov'
					$VideoLen = $_.photo.livePhoto.totalStreamSize - $_.size
                    if ( (Test-Path($TargetPath+$_.name)) -and # Target image exists
                         (Test-Path($TargetPath+$VideoName)) -and # Target video exists
                         ((Get-Item($TargetPath+$_.name)).Length -eq $_.size) -and # size of image is onedrive's size
                         ((Get-Item($TargetPath+$VideoName)).Length -eq $VideoLen) # size of video is onedrive's size
                       ) {
                        Write-Output "Live photo $($_.id) - $($CurrentPath + $_.name) already exists at $($TargetPath) - skipping."
                    } else {
                        Write-Output("Detected live photo $($_.id) - $($CurrentPath + $_.name). Saving image/video pair to $($TargetPath)")
                        Download-SingleLivePhoto -AccessToken $AccessToken -ElementId $_.id -DriveId $DriveId -SaveTo $TargetPath -ImageName $_.name -VideoName $VideoName -ExpImgLen $_.size -ExpVidLen $VideoLen -LastModified $_.fileSystemInfo.lastModifiedDateTime
                    }
                }
            }
        }
    }
    if ([bool]$Response.PSobject.Properties["@odata.nextLink"]) 
    {
        write-debug("Getting more elements form service (@odata.nextLink is present)")
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
    .PARAMETER DriveId
    OneDrive DriveId of a LivePhoto
    .PARAMETER SaveTo
    Target path where to save live photos.
    .PARAMETER ExpImgLen
    Expected length of image file, as reported in containing folder.
    .PARAMETER ExpVidLen
    Expected length of video file, as reported in containing folder.
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
        [string]$DriveId,
        [Parameter(Mandatory=$True)]
        [string]$SaveTo,
        [Parameter(Mandatory=$True)]
		[string]$ImageName,
        [Parameter(Mandatory=$True)]
		[string]$VideoName,
        [Parameter(Mandatory=$True)]
        [int]$ExpImgLen,
        [Parameter(Mandatory=$True)]
        [int]$ExpVidLen,
        [Parameter(Mandatory=$True)]
        [datetime]$LastModified
    )
    
    if (!(Test-Path $SaveTo)) { New-Item -ItemType Directory -Force $SaveTo | Out-Null }
    
    # video part
	if ($DriveId -eq '') {
		$Uri = "https://my.microsoftpersonalcontent.com/_api/v2.1/drive/items/$($ElementId)/content?format=video"
	} else {
		$Uri = "https://my.microsoftpersonalcontent.com/_api/v2.1/drives/$($DriveId)/items/$($ElementId)/content?format=video"
	}
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $FileName = $SaveTo+$VideoName
	If (Test-Path($FileName)) { Remove-Item ($FileName) }
    $WebRequest=Invoke-WebRequest -Method "GET" -Uri $Uri -Header @{ Authorization = "BEARER "+$AccessToken } -ErrorAction SilentlyContinue -OutFile $FileName -PassThru
    $ActualLen = $WebRequest.RawContentLength
    (Get-Item ($FileName)).LastWriteTime = $LastModified

	If ($ActualLen -ne $ExpVidLen)  { Write-Error("Error saving video part of live photo $ElementId. Got $ActualLen bytes, expected $ExpVidLen bytes.") }

    # image part
	if ($DriveId -eq '') {
		$Uri = "https://my.microsoftpersonalcontent.com/_api/v2.1/drive/items/$($ElementId)/content"
	} else {
		$Uri = "https://my.microsoftpersonalcontent.com/_api/v2.1/drives/$($DriveId)/items/$($ElementId)/content"
	}
    Write-Debug("Calling OneDrive API")
    Write-Debug($Uri)
    $FileName = $SaveTo+$ImageName
	If (Test-Path($FileName)) { Remove-Item ($FileName) }
    $WebRequest=Invoke-WebRequest -Method "GET" -Uri $Uri -Header @{ Authorization = "BEARER "+$AccessToken } -ErrorAction SilentlyContinue -OutFile $FileName -PassThru
    $ActualLen = $WebRequest.RawContentLength
    (Get-Item ($FileName)).LastWriteTime = $LastModified
    
    if ($ActualLen -ne $ExpImgLen) { Write-Error("Error saving photo part of live photo $ElementId. Got $ActualLen bytes, expected $ExpImgLen bytes.") }

}


Write-Output "Live Photo downloader - Downloads Live Photos from OneDrive camera roll as saved by OneDrive iOS app."
Write-Output "(C) 2024-2025 Petr Vyskocil. Licensed under MIT license."
Write-Output ""


# This disables powershell progress indicators, speeding up Invoke-WebRequest with big results by a factor of 10 or so
$ProgressPreference = 'SilentlyContinue'

if ($AccessToken -eq '') {
	Write-Output "Getting OneDrive Authentication token..."
	$auth = Get-ODPhotosToken
	$AccessToken = $auth.'AccessToken-OneDrive.ReadWrite'
}

Write-Output "Downloading Live Photos..."
Download-LivePhotosAuth -AccessToken $AccessToken -PathToScan $PathToScan -SaveTo $SaveTo
