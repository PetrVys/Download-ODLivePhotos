# Download-ODLivePhotos

## About
Personal OneDrive supports iOS Live Photos (each photo consisting of a pair of an image file and a quicktime movie, linked by EXIF tags) and backs them up using the iOS OneDrive application. They can then be viewed by the [iOS OneDrive app](https://apps.apple.com/us/app/microsoft-onedrive/id477537958) and on the [Photos web application](https://photos.onedrive.com). but you cannot access them using any desktop client or desktop Windows Photo app. Also, once you copy or move the files in any way, they lose the video part and are turned into a static picture.

This utility authorizes as the [Photos app](https://photos.onedrive.com) to OneDrive API (because Live Photos are not available using the public APIs) and downloads all Live Photos to a target folder.

## Usage
```
PS> .\Download-ODLivePhotos.ps1 -SaveTo 'c:\Live Photos' -PathToScan '\Pictures\Camera Roll\2024'
```
* -SaveTo - Path where to save the Live Photos
* -PathToScan - Path within your Personal OneDrive from where you want to download the Live Photos from

## Troubleshooting
There is no retry implemented in case OneDrive API fails for whatever reason. Unfortunately, when using the tool I see about 0.5% error rate when calling the API. If you see any error or if you're downloading big library, just rerun the command. Files already successfully downloaded are skipped on subsequent runs, so the next run will finish significantly faster.

This utility only supports personal OneDrive accounts - to the best of my knowledge business/SharePoint accounts do not support Live Photos. Even if they did, I have no way to test them...
