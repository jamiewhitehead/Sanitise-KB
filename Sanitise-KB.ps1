<#
.SYNOPSIS
    Remove a KB article from Configuration Manager Infrastructure.
.DESCRIPTION
    Use this script if you need to remove a KB article from being deployed in Configuration Manager. 
    The script will
        - Remove the KB article from any Software Update Groups
        - Remove the KB article content from any Software Update Packages
        - Add\extend a filter rule to any Automatic Deployment Rules that are associated withe the KB, to ensure that it is not re-deployed.
.PARAMETER KB2Remove
    The Article number to be removed. Required format is KB*******
.EXAMPLE
    .\Sanitise-KB -KB2Remove KB3203467
.Credits
    KB removal from Software Update Group and Package heavily influenced by https://www.scconfigmgr.com/2014/11/18/remove-expired-and-superseded-updates-from-a-software-update-group-with-powershell/
    XML manipulation used to create\extend the ADR filters was inspired by https://www.petervanderwoude.nl/post/changing-the-deployment-package-linked-to-an-automatic-deployment-rule-in-configmgr-2012/
.NOTES
    Script name: Sanitise-KB.ps1
    Author:      Jamie Whitehead
    Contact:     jamie.whitehead@dilignet.com
    DateCreated: 07-08-17
#>

PARAM(
    [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
    [ValidateScript({if($($_.length -eq "9") -and $($_.substring(0,2).ToUpper() -eq "KB")){$True} Else {$False}})]
    [String]$KB2Remove
    )
    
Function Write-Log
{
<#
  Purpose: This function writes log files in CMTrace log format.

  Credit
  -------------------
  Name                         Date                 Version          Description
  ---------------------------------------------------------------------------------------------------
  Kaido Järvemets              13.09.11             1.1              http://configmgrpsh.codeplex.com/
  Jamie Whitehead              05.12.13             1.2
  Jamie Whitehead              01.09.15             1.3              Added default parameters.
 
  Severity="0"    CMTrace Log-procedure delete logfile an create new logfile 
  Severity="1"    display as normally line 
  Severity="2"    display as yellow line / warn 
  Severity="3"    display as red line / error 
  Severity="F"    display as red line / error
#>

    PARAM(
         [String]$Message,
         [String]$Path = $LogFileLocation,
         [int]$severity = 1,
         [string]$component = $DefaultLogComponent
         )
         
         $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
         $Date= Get-Date -Format "HH:mm:ss.fff"
         $Date2= Get-Date -Format "MM-dd-yyyy"
         
         "<![LOG[$Message]LOG]!><time=$([char]34)$date+$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath $Path -Append -NoClobber -Encoding default

} # End of Write-Log function

Function Check-CMPrerequistes
{
    # Check that you are not running in X64 powershell
    if ([Environment]::Is64BitProcess -eq $True)
    {    
        Write-Log -Message "Need to run at a X86 PowershellPrompt" -severity 3
        Throw "Need to run at a X86 PowershellPrompt"
    }

    # Load ConfigMgr module if it isn't loaded already 
    if (-not(Get-Module -name ConfigurationManager)) 
    {
        Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
    }
}   # End Function Check-CMPrerequistes


Function Remove-KBFromSUG
{
    PARAM(
         [Array]$FnSU_CIIDArr,
         [String]$FnSMSNamespace,
         [String]$FnSiteServer
         )

    foreach ($SUGItem in Get-CMSoftwareUpdateGroup)
    {
        Write-Log -Message "Checking Software update group $($SUGItem.LocalizedDisplayName)"
        #Let's check to see if updates to be removed are present in the current SUG.
        foreach($SU_CIID in $FnSU_CIIDArr)
        {
        
            if($($SUGItem.Updates) -contains $SU_CIID)
            {
                Write-log -Message "The Software update group $($SUGItem.LocalizedDisplayName) contains the Software update to be removed ($SU_CIID)."
                # Create up the wmi object for the SUG
                $SUG = [wmi](Get-WmiObject -Namespace $FnSMSNamespace -Class SMS_AuthorizationList -ComputerName $FnSiteServer -Filter "LocalizedDisplayName like '$($SUGItem.LocalizedDisplayName)'" -ErrorAction Stop).__PATH
                
                #Create up the new list of updates for the SUG not including the update to be removed.         
                $SUG.Updates = $SUGItem.Updates | Where-Object { $_ -ne $SU_CIID}
                $SUG.Put() | Out-Null
                #Update the SUGItem update list. Otherwise when we go throught the next cycle, the first update that was removed will still be in the listing. 
                $SUGItem.Updates = $SUG.Updates
                Write-log -Message "The Software update has been removed."
            }
        }
    }
} # End Function Remove-KBFromSUG

Function Remove-KBContent
{
    PARAM(
         [Array]$FnSU_CIIDArr,
         [String]$FnSMSNamespace,
         [String]$FnSiteServer
         )


    foreach($SU_CIID in $FnSU_CIIDArr)
    {    
        # Find out where the content information for the KB article
        $ContentData = Get-WmiObject -Namespace $FnSMSNamespace -Query "SELECT SMS_PackageToContent.ContentID,SMS_PackageToContent.PackageID from SMS_PackageToContent JOIN SMS_CIToContent on SMS_CIToContent.ContentID = SMS_PackageToContent.ContentID where SMS_CIToContent.CI_ID in ($($SU_CIID))" -ComputerName $FnSiteServer -ErrorAction Stop

        If($ContentData.length -gt 0)
        {
            foreach ($Content in $ContentData)
            {
                # Determine the ContetntID and Package ID where the content is stored. 
                $ContentID = $Content | Select-Object -ExpandProperty ContentID
                $PackageID = $Content | Select-Object -ExpandProperty PackageID
                $SUP = Get-WmiObject -Namespace $FnSMSNamespace -Class SMS_SoftwareUpdatesPackage -ComputerName $FnSiteServer -Filter "PackageID like '$($PackageID)'" 
    
                # Remove the Contetn from the SUP
                $ReturnValue = $SUP.RemoveContent($ContentID, $false)
                if ($ReturnValue.ReturnValue -eq 0)
                {
                    Write-log -Message "Successfully removed ContentID '$($ContentID)' from PackageID '$($PackageID)'"
                }
            }
        }
        Else
        {
            Write-log -Message "No content found for $KB2Remove"
        }
    }
} # End Function Remove-KBContent

Function Add-KBArticleADRFilter
{
    PARAM(
         [String]$FnKB2Remove,
         [Array]$FnSUProductList,
         [String]$FnSMSNamespace,
         [String]$FnSiteServer
         )

    $ADRKBFilter = "-$($FnKB2Remove.substring(2))"

    Foreach ($ADR in Get-CMSoftwareUpdateAutoDeploymentRule)
    {
        Write-log -Message "Testing ADR - $($ADR.Name)"

        #Get the Deployment rule data from wmi object and convert to an XML object
        [wmi]$AutoDeploymentRule = (Get-WmiObject -Class SMS_AutoDeployment -Namespace $FnSMSNamespace -ComputerName $FnSiteServer | Where-Object -FilterScript {$_.AutoDeploymentID -eq $ADR.AutoDeploymentID}).__PATH
        [xml]$UpdateRuleXML = $AutoDeploymentRule.UpdateRuleXML

        # We only want the ADRs that contain the product sets related to the selected update  .
        If ($($FnSUProductList) -contains $((($UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem | Where-Object {$_.propertyname -eq "_Product"}).MatchRules.string).replace("'","")))
        {
            Write-log -Message "The ADR - $($ADR.Name) contains a product related to $KB2Remove. Processing."
        
            If ($($UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem | Where-Object {$_.propertyname -eq "ArticleID"}) -eq $null)
            {
                write-log -Message " An existing Article rule does not exist, creating one."
    
                #Create the filter rule by cloning an existing the first XML element (does not matter what element is first). Modify the cloned element settings to specify the filter. Finally, add as a new child element. 
                $element = $UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem[0].clone()
                $element.PropertyName="ArticleID"
                $element.MatchRules.string = $ADRKBFilter
                $UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.AppendChild($element)
            }
            Else
            {
                write-log -Message "An existing Article rule exists. Appending the existing rule with the filter $ADRKBFilter."
                     
                # update the existing filter by cloning a rule XML node, update the value and add as a new child element.
                # Get the Node to be cloned in xml object format (Powershell tends to treat single value text elements as strings...)
                $NodeToCLone = ($UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem | Where-Object {$_.propertyname -eq "ArticleID"}).MatchRules.SelectSingleNode('//string').clone()
            
                #Update the value in the cloned element
                $NodeToCLone.'#text' = $ADRKBFilter
            
                #Append the cloned element
                ($UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.UpdateXMLDescriptionItem | Where-Object {$_.propertyname -eq "ArticleID"}).MatchRules.AppendChild($NodeToCLone) | Out-Null
            }

            #Write the XMLObject to the Automatic Deployment rule object and then commit the object to Configuration Manager. 
            $AutoDeploymentRule.UpdateRuleXML = $UpdateRuleXML.OuterXML
            $AutoDeploymentRule.Put() | Out-Null
        }
        Else
        {
            Write-log -Message "The ADR - $($ADR.Name) does not contain a product rule related to $KB2Remove."
        }
    }
} # End Function Add-KBArticleADRFilter


##################
### Begin MAIN code
###################

# Initialise variables
$LogFileLocation = $PSCommandPath -replace ".ps1",".log"
$DefaultLogComponent = "Sanitise $KB2Remove"
$CurrentLocation = Split-Path -Parent $PSCommandPath

$SiteServer = $env:computername
$SiteCode = (get-WMIObject -ComputerName $SiteServer -Namespace "root\SMS" -Class "SMS_ProviderLocation" | where-object {$_.ProviderForLocalSite -eq $true}).sitecode
$SMSNamespace = "root\SMS\site_$SiteCode"

$SU_CIIDArr = @()
$SUProductList = @()

Write-Log -Message "=================================================================="
Write-Log -Message "Start of Log for : $(Split-Path -Leaf $PSCommandPath)"             
Write-Log -Message "==================================================================" 
Write-Log -Message ""

Check-CMPrerequistes

# Set the PS Drive location to run the CM PS cmdlets.
Set-Location ${SiteCode}:

#Suppress fast switch warning
$CMPSSuppressFastNotUsedCheck = $true



#get CI_IDs of the updates, note that for each KB article there may be multiple updates (e.g. x64\x32). We also get the associated products (for the ADR filters) here. 

$SUList = Get-CMSoftwareUpdate -Fast | Where-Object {$_.LocalizedDisplayName -like "*$($KB2Remove)*"}
Foreach ($SU in $SUList)
{
    Write-log -Message "Adding $($SU.LocalizedDisplayName) to the Software update list."
    #$SU_CIIDArr += (Get-CMSoftwareUpdate -Name $($SU.LocalizedDisplayName)).CI_ID
    $SU_CIIDArr += $SU.CI_ID
    $SUProductList += $SU.CategoryInstance_UniqueIDs | Where-Object {$_ -like "Product*" -and  $SUProductList -notcontains $_}
}


# Remove the updates from all Software Update Groups (SUGs)
Remove-KBFromSUG -FnSU_CIIDArr $SU_CIIDArr -FnSiteServer $SiteServer -FnSMSNamespace $SMSNamespace

# Remove the content associated with the KBs from the Software update Packages (SUPs)
Remove-KBContent -FnSU_CIIDArr $SU_CIIDArr -FnSiteServer $SiteServer -FnSMSNamespace $SMSNamespace


# Create the ADR Article filter so that the KB does not get re-added. 
Add-KBArticleADRFilter -FnKB2Remove $KB2Remove -FnSiteServer $SiteServer -FnSMSNamespace $SMSNamespace -FnSUProductList $SUProductList

#Return to the location that the script was called from. 
Set-Location $CurrentLocation