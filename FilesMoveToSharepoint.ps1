# afm\orders
#S-ABF1-MAN111
# read all data from folder. 
#################################
# Move local files from folde AFM\Orders to sharepoint connected to a customer asset.
# Jeroen Beemster Spring 2024
#################################
<#
$pathArray = @(
    "C:\Orders\3175*", 
    "C:\Orders\31*", 
    "C:\Orders\317*", 
    "C:\Orders\3*"
)
$pathArray = @(
    "N:\Orders\0*", 
    "N:\Orders\1*", 
    "N:\Orders\2*", 
    "N:\Orders\3*", zwaag
    "N:\Orders\4*", done
    "N:\Orders\5*", yuku
    "N:\Orders\6*", amcas 60000 - 60250 2017
    "N:\Orders\7*", 
    "N:\Orders\8*", 
    "N:\Orders\9*" apac

    )
    
    #>

# Load the file and convert it into a PowerShell object
$keys = Get-Content -Path "keys.json" | ConvertFrom-Json
    
# Access your keys
[string] $oAuthTokenEndpoint = $keys.oAuthTokenEndpoint
[string] $appId = $keys.appId
[string] $clientSecret = $keys.clientSecret
[string] $SharepointClientID = $keys.SharepointClientID
[string] $SharepointTenant = $keys.SharepointTenant
[string] $customerAssetSharepointLocation = $keys.customerAssetSharepointLocation
[string] $SharepointThumbPrint = $keys.SharepointThumbPrint
[string] $dataverseEnvUrl = $keys.dataverseEnvUrl
[string] $sharepointSite = $keys.sharepointSite


Write-Host "The Tenant we are working with is: $SharepointTenant"


    

##########################################################
# Access Token Request
##########################################################
    
# OAuth Body Access Token Request
$authBody = 
@{
    client_id     = $appId;
    client_secret = $clientSecret;    
    scope         = "$($dataverseEnvUrl)/.default"    
    grant_type    = 'client_credentials'
}

# Parameters for OAuth Access Token Request
$authParams = 
@{
    URI         = $oAuthTokenEndpoint
    Method      = 'POST'
    ContentType = 'application/x-www-form-urlencoded'
    Body        = $authBody
}

# Get Access Token


function GetNextLinkFromDataverse {
    param (
        $URI 
    )
                    
    $apiCallParams =
    @{
        URI     = $URI 
    
        Headers = @{
            "Authorization" = "$($authResponse.token_type) $($authResponse.access_token)"
            'Prefer'        = 'odata.include-annotations="*"'
        }
        Method  = 'GET'
    }
        
    $apiCallRequest = Invoke-RestMethod @apiCallParams -ErrorAction Stop
    return $apiCallRequest
}
    
function GetDataFromDataverse {
    param (
        $UriParams 
    )
    $authResponse = Invoke-RestMethod @authParams -ErrorAction Stop
                    
    $apiCallParams =
    @{
        URI     = "$($dataverseEnvUrl)/api/data/v9.2/$($uriParams)"
        Headers = @{
            "Authorization" = "$($authResponse.token_type) $($authResponse.access_token)"
            'Prefer'        = 'odata.include-annotations="*"'
        }
        Method  = 'GET'
    }
    $apiCallRequest = Invoke-RestMethod @apiCallParams -ErrorAction Stop
        
    #check for over 5000 records
    $nextData = $apiCallRequest
    while ($null -ne $nextData."@odata.nextLink") {    
        $nextData = GetNextLinkFromDataverse($nextData."@odata.nextLink")
        $apiCallRequest.value += $nextData.value 
    }  
    return $apiCallRequest
}
    
    
##########################################################
# Get files from folder
##########################################################

function getFilesInFolder($path) {
    Write-host "Get Files in folder" $path 
    $filesInFolder = Get-ChildItem -Path "$path" -Force -Recurse -include "*.jpg*", "*.doc*", "*.rtf*", "*.pdf*" | Select-Object Length, DirectoryName, name, Mode
    
    #Write-Host $filesInFolder.Count 
    $filesInFolder = ($filesInFolder | Where-object { $_.Mode -ne "d----" }) # remove folders that has extention like .loc
    $filesInFolder = ($filesInFolder | Where-object { $_.Mode -ne "larhs" }) # remove files that are labelled as systemfiles, directory name is not filled due to that 
    $filesInFolder = ($filesInFolder | Where-object { $_.name.ToUpper().EndsWith(".LNK") -eq $False }) # remove shortcuts  


    $totalFileSize = $filesInFolder |  measure-object -property length -sum 
    $totalFileSize = [math]::round($totalFileSize.sum / 1Mb, 1)
    Write-Host 'Files in folder   '  $filesInFolder.Count "With total Size of: " $totalFileSize "Mb" -f Green
    #add assetnumber
    $filesInFolder | Add-Member -MemberType NoteProperty -Name 'AssetNr' -Value ""
    $filesInFolder | Add-Member -MemberType NoteProperty -Name 'ProjectNr' -Value ""
    $filesInFolder | Add-Member -MemberType NoteProperty -Name 'ExistInCRM' -Value $False
    $filesInFolder | ForEach-Object {
        $firstSlash = $_.DirectoryName.IndexOf("\")
        $secondSlash = $firstSlash + 1 + $_.DirectoryName.substring($firstSlash + 1).IndexOf("\")
        $thirdSlash = $secondSlash + 1 + $_.DirectoryName.substring($secondSlash + 1).IndexOf("\")
        $fourthSlash = $thirdSlash + 1 + $_.DirectoryName.substring($thirdSlash + 1).IndexOf("\")
        #Write-Host $firstSlash $secondSlash  $thirdSlash $fourthSlash $_.DirectoryName -f Blue
        $assetNr = $null
        if (($thirdSlash -ne -1) -and ($_.DirectoryName.substring($thirdSlash).length -ge 6)) {
            $_.projectNr = $_.DirectoryName.substring($thirdSlash + 1, 5)
            if ($fourthSlash -ne -1) { 
                if ($_.DirectoryName.length -ge ($fourthSlash + 1 + 8) ) {
                    # when there is no fourth slash
                    $assetNr = $_.DirectoryName.substring($fourthSlash + 1, 8)
                }
            }
        }
            
        $_.AssetNr = $_.projectNr
     
        if (($null -ne $assetNr ) -and ($_.projectNr -eq ($assetNr.Substring(0, 5)))) {
            $_.AssetNr = $assetNr
        }

    }

    Write-Host 'Start labeling' $path

    $filesInFolder | Add-Member -MemberType NoteProperty -Name 'label' -Value ""

    $filesInFolder | ForEach-Object {
        #Write-Host "X" $_.DirectoryName $_.AssetNr $_.projectNr $_.name "X"
        if ($_.DirectoryName.ToUpper().Contains($_.AssetNr + "\PICTURES")  ) {
            $_.label = 'Pictures' 
        }

        #if ($_.DirectoryName.EndsWith($_.projectNr) -and $_.name.ToUpper().EndsWith(".JPG") ) {
        #    $_.label = 'Pictures' 
        #}

        if ($_.name.ToUpper().EndsWith(".JPG") ) {
            $_.label = 'Pictures' 
        }



        if ($_.DirectoryName.ToUpper().EndsWith("MANUAL") -and $_.name.ToUpper().EndsWith(".PDF")) {
            # ignore subfolders including 'Obsolete' subfolder
            # ignores all docx and  
            $_.label = 'Manuals' 
        }
    
        # not only folder, include files that starts with manual 
        if ($_.DirectoryName.EndsWith($_.projectNr) -and $_.name.ToUpper().StartsWith("MANUAL") -and $_.name.ToUpper().EndsWith(".PDF")) {  
            $_.label = 'Manuals' 
        }       

        #if ($_.DirectoryName.EndsWith($_.projectNr) -and { 
        if ($_.name.ToUpper().StartsWith("MANUAL") -and $_.name.ToUpper().EndsWith(".PDF")) {  
            Write-Host $_.name, $_.projectNr, $_.DirectoryName
        }       


        if ($_.DirectoryName.EndsWith($_.projectNr) -and ($_.name.StartsWith("SP")  )) {
            $_.label = 'Spares' 
        }

        if ($_.DirectoryName.EndsWith($_.projectNr) -and ($_.name.ToUpper().StartsWith("TECHNICAL")  )) {
            $_.label = 'Specifications' 
        }
        if ($_.DirectoryName.EndsWith($_.projectNr) -and ($_.name.StartsWith("TS")  )) {
            $_.label = 'Specifications' 
        }       
        if ($_.DirectoryName.EndsWith($_.projectNr) -and ($_.name.StartsWith("VO")  )) {
            $_.label = 'Specifications' 
        }

            
        if ($_.DirectoryName.ToUpper().EndsWith($_.AssetNr + "\DRAWINGS")  ) {
            # ignore subfolders including 'Obsolete' subfolder
            $_.label = 'Drawings' 
        }
        if ($_.DirectoryName.ToUpper().Contains($_.AssetNr + "\AANDRIJVING")  ) {
            $_.label = 'Drives' 
        }
        if ($_.DirectoryName.ToUpper().Contains($_.AssetNr + "\DRIVES")  ) {
            $_.label = 'Drives' 
        }

    }

    #remove not labeled files
    $filesInFolder = ($filesInFolder | Where-object { $_.label -ne "" })

    # start fill in $assetNr after labeling 

    $filesInFolder | ForEach-Object {

        if ($_.assetNr -eq $_.ProjectNr) {
            if ($_.label -in 'Manuals', 'Spares', 'Specifications') {
                $projectNumberStartingPoint = $_.Name.IndexOf($_.ProjectNr)
                if (($null -ne $projectNumberStartingPoint) -and ($_.Name.length -gt ($projectNumberStartingPoint + 5 + 3))) {
                    $assetNr = $_.ProjectNr + $_.Name.substring($projectNumberStartingPoint + 5, 3)  
                    $assetNr = $assetNr.Replace("_", "-")
                    $assetNr = $assetNr.Replace(".", "-")

                    if ($assetNr.Replace("-", "") -match "^\d+$") {
                        #is not a number
                        $_.assetNr = $assetNr
                    }
                }
            }


        }
    }

    $fileSize = ($filesInFolder |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    #Write-Host "Files labeled " $filesInFolder.Count -f Red
    return $filesInFolder

}

##########################################################
# Find Asset in Database
##########################################################

function getuniqueAssets($filesInFolder, $assetGroup) {

    Write-Host "running" 
    $uniqueAssets = $filesInFolder | Select-Object AssetNr -uniq # select distinct items
    $uniqueAssets | Add-Member -MemberType NoteProperty -Name 'ExistInCRM' -Value $False
    $uniqueAssets | Add-Member -MemberType NoteProperty -Name 'FirstAssetNumber' -Value $null
    $uniqueAssets | Add-Member -MemberType NoteProperty -Name 'ProjectNr' -Value $null
        

    #$selectString = 'msdyn_customerassets?$filter=startswith(amb_serialnumber, ''' + $assetGroup + ''')' +
    #' and (statecode eq 0) ' + 
    #'&$select=amb_serialnumber'

    $selectstring = 'msdyn_customerassets?$filter=startswith(amb_serialnumber, ''' + $assetgroup + ''')' + 
    ' and (statecode eq 0) ' + 
    '&$select=amb_serialnumber' + 
    '&$orderby=amb_serialnumber asc'


    $customerAssets = GetDataFromDataverse($selectString)
    if ($customerAssets.value.amb_serialnumber.Count -lt 1) {
        Write-Host "No Serialnumbers found"
        break
        return
    } 

    $uniqueAssets | ForEach-Object {
        $_.existInCRM = $customerAssets.value.amb_serialnumber.Contains($_.AssetNr) 
   
        if ($_.ExistInCRM -eq $False) {
            $AssetNr = $_.AssetNr
            # most likely projectnumber, find first most near by
            $foundCustomerAssets = $customerAssets.value.amb_serialnumber.Where({ $_ -like $AssetNr + "*" })

            if ($foundCustomerAssets -gt 0) {
                $_.ExistInCRM = $true
                $_.ProjectNr = $_.AssetNr
                $_.AssetNr = $foundCustomerAssets[0] 
            }    
        }
    }

    #remove not existing Assets
    $uniqueAssets = ($uniqueAssets | Where-object { $_.ExistInCRM })

    # set existsincrm per file AND when add assetnumber first of project when only projectnumber exists. 
    if ($null -ne $uniqueAssets) {
        $filesInFolder | ForEach-Object { 
            $_.ExistInCRM = ($_.AssetNr -in $uniqueAssets.AssetNr)
            $index = [Array]::IndexOf($uniqueAssets.ProjectNr, $_.AssetNr )
            if ($index -ne -1) {
                if ($null -ne $uniqueAssets[$index].ProjectNr ) {
                    $_.AssetNr = $uniqueAssets[$index].AssetNr
                    $_.ExistInCRM = $true
                }
            }
        }
    }
    #remove duplicate assetnumbers (to be able to create subfolder once)
    $uniqueAssets = $uniqueAssets | Select-Object AssetNr -uniq # select distinct items
    return $filesInFolder, $uniqueAssets
}


##########################################################
# write to file
##########################################################
function exportFilesToDocument($filesInFolder) {
    $csvfileName = ".\Files_" + $path.Replace("\", "_").Replace(":", "_").Replace("*", "Star") + ".csv"

    if ($filesInFolder.Count -ne 0) {
        $filesInFolder | Export-Csv -Path $csvfileName -NoTypeInformation
    }
    ##########################################################
    # list on screen
    ##########################################################

    #Write-Host ( "{0,6} | {1,9} | {2,15} | {3,8} | {4,-30} | {5,-40}" -f "InCRM", "AssetNr", "label", "length", "name", "DirectoryName")          

    $filesInFolder | ForEach-Object {
        #Write-Host ( "{0,6} | {1,9} | {2,15} | {3,8} | {4,-30} | {5,-40}" -f $_.ExistInCRM, $_.AssetNr, $_.label, $_.length, $_.name, $_.DirectoryName)          
    }
        
    ##########################################################
    # Show Totals
    ##########################################################
    $txtfileName = ".\Files_" + $path.Replace("\", "_").Replace(":", "_").Replace("*", "Star") + ".txt"
    Start-Transcript -Path $txtfileName -UseMinimalHeader
    Write-Host 'Files labeled     ' $filesInFolder.Count "With total Size of: " $fileSize "Mb" -f Green
    #$filesInFolder | Format-Table
    $filesCount = ($filesInFolder | Where-object { $_.label -match "Pictures" }).Count
    $fileSize = ($filesInFolder | where-object { $_.label -match "Pictures" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with picture' $filesCount "With total Size of: " $fileSize "Mb" -f Green 


    $filesCount = ($filesInFolder | Where-object { $_.label -match "Manuals" }).Count
    $fileSize = ($filesInFolder | Where-object { $_.label -match "Manuals" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with manuals ' $filesCount "With total Size of: " $fileSize "Mb" -f Green

    $filesCount = ($filesInFolder | Where-object { $_.label -match "Spares" }).Count
    $fileSize = ($filesInFolder | Where-object { $_.label -match "Spares" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with spares   ' $filesCount "With total Size of: " $fileSize "Mb" -f Green

    $filesCount = ($filesInFolder | Where-object { $_.label -match "Specifications" }).Count
    $fileSize = ($filesInFolder | Where-object { $_.label -match "Specifications" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with TS       ' $filesCount "With total Size of: " $fileSize "Mb" -f Green

    $filesCount = ($filesInFolder | Where-object { $_.label -match "Drawings" }).Count
    $fileSize = ($filesInFolder | where-object { $_.label -match "Drawings" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with drawings' $filesCount "With total Size of: " $fileSize "Mb" -f Green 

    $filesCount = ($filesInFolder | Where-object { $_.label -match "Drives" }).Count
    $fileSize = ($filesInFolder | where-object { $_.label -match "Drives" } |  measure-object -property length -sum) 
    $fileSize = [math]::round($fileSize.sum / 1Mb, 2)
    Write-Host 'Files with drives' $filesCount "With total Size of: " $fileSize "Mb" -f Green 


    Stop-Transcript
}
#*****************************
# Get connection to sharepoint
#*****************************
function createNewStructureIfNotExist() {
    param (
        [string] $location,
        [string] $newFolder
    )
        
    Try {

        Write-Host 'Create -Folder' $newFolder 'in' $location
        Add-PnPFolder -Name $newFolder -Folder $location
    }
    Catch {
        #if exist than OK otherwise show message
        if ($_.FullyQualifiedErrorId -ne 'InvalidOperation,PnP.PowerShell.Commands.Files.AddFolder') {
            write-host -f Red "Error:" $_.Exception.Message
            write-host -f Red "Error:" $_.FullyQualifiedErrorId
            $_
            break  
        }
    }
        
}

#*********************************************
# Move files from local folder into sharepoint
#*********************************************

function moveFilesFromLocalFolderIntoSharepoint($filesInFolder, $location) {

    $filesInFolder | ForEach-Object {
        Try {

            if ($_.ExistInCRM) {
                $fullFileNameFrom = $_.DirectoryName + '\' + $_.Name 
                $fullFileNameTo = $location + '/' + $_.assetNr + '/Managed Documents/' + $_.label
                Write-Host 'From' $fullFileNameFrom -ForegroundColor green
                Write-Host 'To  ' $fullFileNameTo  -ForegroundColor green
                Add-PnPFile -Path $fullFileNameFrom -Folder $fullFileNameTo
                
            }    
        }
        Catch {
            write-host -f Red "Error:" $_.Exception.Message
            write-host -f Red "Error:" $_.FullyQualifiedErrorId
            $_
            break  
        }
 
    }

}

function createFolderStructurePerAsset($uniqueAssets) {
    $SharepointConn = Connect-PnPOnline -Url $sharepointSite -ClientId $SharepointClientID -Thumbprint $SharepointThumbPrint -Tenant $SharepointTenant 
    $SharepointConn
    $uniqueAssets | ForEach-Object {
        $baseLoc = $customerAssetSharepointLocation + '/' + $_.assetNr + '/Managed Documents' 
        createNewStructureIfNotExist $baseLoc 'Pictures'
        createNewStructureIfNotExist $baseLoc 'Manuals'
        createNewStructureIfNotExist $baseLoc 'Spares'
        createNewStructureIfNotExist $baseLoc 'Specifications'
        createNewStructureIfNotExist $baseLoc 'Drawings'
        createNewStructureIfNotExist $baseLoc 'Drives'

    }
}


#*********************************************
# Main Routine
#*********************************************

function mainRoutine {
    param (
        $path, 
        $assetGroup
    )
    $globalFilesInFolder = getFilesInFolder($path)
    if ($globalFilesInFolder.Count -gt 0) { 
             
        $globalFilesInFolder, $globalUniqueAssets = getuniqueAssets $globalFilesInFolder $assetGroup #slow -> duegetCRMlink
        exportFilesToDocument($globalFilesInFolder)
        #createFolderStructurePerAsset($globalUniqueAssets)
        #moveFilesFromLocalFolderIntoSharepoint $globalFilesInFolder $customerAssetSharepointLocation 
    }
}

$SharepointConn = Connect-PnPOnline -Url $sharepointSite -ClientId $SharepointClientID -Thumbprint $SharepointThumbPrint -Tenant $SharepointTenant 
$SharepointConn

#mainRoutine "N:\Orders\6*" "6"
#mainRoutine "N:\Orders\9*" "9"
#mainRoutine "N:\Orders\30000-30249\30013" "3"
#mainRoutine "N:\Orders\30000-30249" "3"
#mainRoutine "C:\Orders\3*" "3"

#mainRoutine "N:\Orders\60000-60249" "6"
#mainRoutine "N:\Orders\60000-60249\60079" "6"
mainRoutine "C:\Orders\60000-60249" "6"
