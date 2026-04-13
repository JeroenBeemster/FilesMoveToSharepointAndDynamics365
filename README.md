## Installation guide:

rename keys.json.example to keys.json
fill in the right credentials.

## run powershell to test

comment out with hashtag the following lines to create .txt and .csv files without any copy. the following 2 lines

```powershell
#createFolderStructurePerAsset($globalUniqueAssets)
#moveFilesFromLocalFolderIntoSharepoint $globalFilesInFolder $customerAssetSharepointLocation
```

for every folder (or group of folders) edit line mainroutine.

```powershell

mainRoutine "C:\Orders\3\*" "3"
```

In powershell run

```powershell
FilesMoveToSharepoint.ps1
```
