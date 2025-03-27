# Download-ODLivePhotos

## About
Personal OneDrive supports iOS Live Photos (each photo consisting of a pair of an image file and a quicktime movie, linked by EXIF tags) and backs them up using the iOS OneDrive application. They can then be viewed by the [iOS OneDrive app](https://apps.apple.com/us/app/microsoft-onedrive/id477537958) and on the [Photos web application](https://photos.onedrive.com). but you cannot access them using any desktop client or desktop Windows Photo app. Also, once you copy or move the files in any way, they lose the video part and are turned into a static picture.

This utility authorizes as the [Photos app](https://photos.onedrive.com) to OneDrive API (because Live Photos are not available using the public APIs) and downloads all Live Photos to a target folder.

## Usage
```
PS> .\Download-ODLivePhotos2.ps1 -SaveTo 'c:\Live Photos' -PathToScan '\Pictures\Camera Roll\2024'
```
* -SaveTo - Path where to save the Live Photos
* -PathToScan - Path within your Personal OneDrive from where you want to download the Live Photos from
* -AccessToken - authentication token to use in the form `BEARER EwBIB......`, "stolen" from onedrive website communication via Dev Tools window (not mandatory, will be obtained automatically if not provided)

## Troubleshooting
The script uses directory %TEMP%\ODLivePhotos for temporary files related to WebView2 component. If you have issues with authentication, delete this directory and try again.

There is no retry implemented in case OneDrive API fails for whatever reason. If you see any error or if you're downloading big library, just rerun the command. Files already successfully downloaded and with correct size are skipped on subsequent runs, so the next run will finish significantly faster.

This utility only supports personal OneDrive accounts - to the best of my knowledge business accounts do not support Live Photos. Given that the personal account backbone is now clearly being migrated to a SharePoint instance (even though access to SharePoint API is closely guarded and has to be obtained in shady ways as you can see in the script), maybe business OneDrive will start supporting Live Photos. Ideally with official and documented API.

There are now two versions of the script, as apparently Microsoft is migrating OneDrive personal users to SharePoint one by one. Since there are significant differences between API and all the accounts should be migrated soon, I did not spend time to create one universal script. Just try which version works for you, the other will throw a bunch of errors and won't work.
