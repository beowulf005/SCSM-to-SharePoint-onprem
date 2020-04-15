Import-Module 'E:\SCAttachments\Microsoft.SharePoint.Client.dll'
Import-Module 'E:\SCAttachments\Microsoft.SharePoint.Client.Runtime.dll'
Import-Module smlets 

#variables
#Ticket Types to process
#[string[]]$TicketTypes = @('System.WorkItem.Problem$','System.WorkItem.ChangeRequest$','System.WorkItem.ReleaseRecord$','System.WorkItem.Incident$','System.WorkItem.ServiceRequest$')  

$Type = “Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectCriteria”

# Get Relationship WorkItem Has File Attachment 
$ItemHasFileAttachment = Get-SCSMRelationshipClass System.WorkItemHasFileAttachment 

$Filters = @()
$date=(get-date).adddays(-5)

# Get Incidents that are Resolved 
$Class = Get-SCSMClass -Name System.WorkItem.Incident$                                              
$Resolved = (Get-SCSMEnumeration -Name “IncidentStatusEnum.Resolved”).ID
$Criteria = “Status = '$Resolved' and ResolvedDate > '$date'”
$Filter = New-Object -Type $Type $Criteria,$Class

$Filters += $Filter

#Get Change Requests that are completed
$Class = Get-SCSMClass -Name System.WorkItem.ChangeRequest$
$Completed = (Get-SCSMEnumeration -Name “ChangeStatusEnum.Completed”).ID
$Criteria = “Status = '$Completed'”
$Filter = New-Object -Type $Type $Criteria,$Class

$Filters += $Filter

#Get Problems that are Closed
$Class = Get-SCSMClass -Name System.WorkItem.Problem$
$Closed = (Get-SCSMEnumeration -Name “ProblemStatusEnum.Closed”).ID
$Criteria = “Status = '$Closed' and ClosedDate > '$date'”
$Filter = New-Object -Type $Type $Criteria,$Class

$Filters += $Filter

#Get Releases that are Closed
$Class = Get-SCSMClass -Name System.WorkItem.ReleaseRecord$
$Completed = (Get-SCSMEnumeration -Name “ReleaseStatusEnum.Completed”).ID
$Criteria = “Status = '$Completed'”
$Filter = New-Object -Type $Type $Criteria,$Class

$Filters += $Filter

#Get Service Requests that are Closed
$Class = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
$Completed = (Get-SCSMEnumeration -Name “ServiceRequestStatusEnum.Completed”).ID
$Criteria = “Status = '$Completed'”
$Filter = New-Object -Type $Type $Criteria,$Class

$Filters += $Filter

# Clear Variable Output 
$Output = '' 
 
# Get all Incidents 
#$Incidents = Get-SCSMObject -Class $IncidentClass 
$TempDir = 'E:\SCAttachments\Temp'
#$TempPropertyName = 'AttachmentTempPath'
$FilesOnItem = 0
$FilesInFolder = 0
$Folder = ""

#functions

function Replace-SpecialChars {
    param($InputString)

    $SpecialChars = '[#?\{\[\(\)\]\}]'
    $Replacement  = ''

    $InputString -replace $SpecialChars,$Replacement
}

function write-out-attachments
{
    param ( $attachmentObjects, $archiveLocation )
    Try
    {
        $seqnum = 0
        $attachmentExportPaths = @()
        $FilesOnItem = 0
        $FilesInFolder = 0

        Foreach ($attachment in $attachmentObjects)
        {           
            Try
            {
                if($Attachment.DisplayName -like "*..*") { $Attachment.DisplayName = $Attachment.DisplayName -replace "..msg",".msg" }
 
                if($Attachment.DisplayName -like "*.cer") { $Attachment.DisplayName = $Attachment.DisplayName -replace ".cer",".xcer" }

                if($Attachment.DisplayName -like "*..*") { $Attachment.DisplayName = $Attachment.DisplayName -replace '\.\.*','.' }
                            
                if($Attachment.DisplayName -like "*&*") { $Attachment.DisplayName = $Attachment.DisplayName -replace "&","and" }

                if($Attachment.DisplayName -like "*:*") { $Attachment.DisplayName = $Attachment.DisplayName -replace ":"," " }

                if($Attachment.DisplayName -like "*\*") { $Attachment.DisplayName = $Attachment.DisplayName -replace "\\","-" }

                if($Attachment.DisplayName -like "*/*") { $Attachment.DisplayName = $Attachment.DisplayName -replace "\/","-" }

                if($Attachment.DisplayName -like "*~*") { $Attachment.DisplayName = $Attachment.DisplayName -replace "~","-" }

                if($Attachment.DisplayName -like "*:*") { $Attachment.DisplayName = $Attachment.DisplayName -replace ":"," " }

                if($Attachment.DisplayName -match '\*') { $Attachment.DisplayName = $Attachment.DisplayName -replace '\*','' }

                if($Attachment.DisplayName -match '#') { $Attachment.DisplayName = $Attachment.DisplayName -replace '#','' }

                if($Attachment.DisplayName -match '%') { $Attachment.DisplayName = $Attachment.DisplayName -replace '%','' }

                if($Attachment.DisplayName -match '\[.+\]' -or $Attachment.DisplayName -match '\{.+\}' -or $Attachment.DisplayName -match '\(.+\)' ) {
                     #$Attachment.DisplayName = $Attachment.DisplayName -replace '\[.+\]',''
                     $Attachment.DisplayName = Replace-SpecialChars -InputString $Attachment.DisplayName 
                }

                if($Attachment.DisplayName.Length -gt 124) {
                    $Attachment.DisplayName = $Attachment.DisplayName.remove(120, ($Attachment.DisplayName.length - 4) - 120) }

                If ($attachmentExportPaths -contains $archiveLocation + $attachment.DisplayName)
                {
                    $seqnum++
                    if ($attachment.Extension.Length -gt 0 ) {
                        $fs = [IO.File]::OpenWrite(($archiveLocation + $attachment.DisplayName.Replace($attachment.Extension, '_' + $seqnum + $attachment.Extension)))
                    } else { 
                        $fs = [IO.File]::OpenWrite($archiveLocation + $attachment.DisplayName + '_' + $seqnum)
                    }
                }
                Else
                {
                    $fs = [IO.File]::OpenWrite(($archiveLocation + $attachment.DisplayName))
                }  

                $memoryStream = New-Object IO.MemoryStream
                $buffer = New-Object byte[] 8192
                [int]$bytesRead|Out-Null
                while (($bytesRead = $attachment.Content.Read($buffer, 0, $buffer.Length)) -gt 0)
                {
                    $memoryStream.Write($buffer, 0, $bytesRead)
                }
                $memoryStream.WriteTo($fs)
            }
            Catch
            {
                $_ | Out-File -FilePath 'E:\SCAttachments\Error\error.txt'-Append
            }
            Finally
            {
                $fs.Close()
                $memoryStream.Close()
                $attachmentExportPaths += $fs.name
            }
        }

        $FilesOnItem = $attachmentobjects.count
        $FilesInFolder = Get-ChildItem -Path $archiveLocation -Recurse -force | Measure-Object
            
        return ,$attachmentExportPaths
    }
    Catch
    {
        $_ | Out-File -FilePath 'E:\SCAttachments\Error\error.txt' -Append
        return
    }
}

# For each Incident found 
 $SharePointURL = "https://sharepoint.documents.local"
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointURL)

Foreach ($Filter in $Filters) 
{ 
    # Get Resolved Incidents
    $Items = Get-SCSMObject -Criteria $Filter
    
    Write-Host $Filter
    Write-Host $Items.Count

    Foreach ($Item in $Items) 
    {
        Try {
            # Get related FileAttachment 
            $FileAttachments = Get-SCSMRelatedObject -SMObject $Item -Relationship $ItemHasFileAttachment  
            # If FileAttachment exists 
            If ($FileAttachments.Count -gt 0) 
            { 
                # Create list of Items with FileAttachment 
                #$Output = $Output + "`r`n`r`n" + $Item.ID + ' - ' + $FileAttachments.Count 

                #create a var to hold the archive folder name for this specific work item
                $ItemTempDir = $TempDir + '\' + $Item.id + '\'
 
                #$Item.TierQueue.Name
                #create an individual folder for each work item
                If (!(Test-Path -PathType Container -Path $ItemTempDir))
                {
                    New-Item -ItemType Directory -Force -Path $ItemTempDir | Out-Null
                }

                #dump attachments to the new dir
                #$errorCatch = New-Object System.Collections.ArrayList
                $errorCatch = write-out-attachments -attachmentObjects $FileAttachments -archiveLocation $ItemTempDir 

                $List = $Ctx.Web.Lists.GetByTitle("ServiceManager")
        
                Try {
                    If ($Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]")
                    {
                        $List = $Ctx.Web.Lists.GetByTitle("Special1")
                    }
                } Catch {}

                Try {
                    If ($Item.SupportGroup.Name.ToString() -like "Enum.[]" -or $Item.SupportGroup.Name.ToString() -eq "Enum.[]" -or $Item.SupportGroup.Name.ToString() -eq "Enum.[]")
                    {
                        $List = $Ctx.Web.Lists.GetByTitle("Special2")
                    }
                } Catch {}

                Try {
                    If ($Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]" -or $Item.TierQueue.Name -like "Enum.[]")
                    {
                        $List = $Ctx.Web.Lists.GetByTitle("Special2")
                    }
                } Catch {}

                Try {
                    If ($Item.TierQueue.Name -like "Enum.[]")
                    {
                        $List = $Ctx.Web.Lists.GetByTitle("Special3")
                    }
                } Catch {}
        
                Try {
                    If ($Item.SupportGroup.Name -eq "Enum.[]")
                    {
                        $List = $Ctx.Web.Lists.GetByTitle("Special4")
                    }
                } Catch {}

        
                $Ctx.Load($List)
                $Ctx.ExecuteQuery()
                Try
                {    
                    $newFolderInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                    $newFolderInfo.UnderlyingObjectType = [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder
                    $newFolderInfo.LeafName = $Item.ID

                    $Folder = $List.AddItem($newFolderInfo)
			        $Folder.Update()
                    $Ctx.Load($List)
                    $Ctx.ExecuteQuery()
                }
                Catch{}

                $Ctx.Load($List)
                $Ctx.Load($List.RootFolder)
                $Ctx.ExecuteQuery()

                $TargetFolder = $Ctx.Web.GetFolderByServerRelativeUrl($List.RootFolder.ServerRelativeUrl + "/" + $Item.ID);
            
                for ($i = 0; $i -lt $errorCatch.Count ; $i++) {
                    $File=Get-Item $errorCatch[$i]
                    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                    $FileCreationInfo.Overwrite = $true
                    $FileCreationInfo.Content = [System.IO.File]::ReadAllBytes($File.FullName)
                    $FileCreationInfo.URL = $File.Name
                    $UploadFile = $TargetFolder.Files.Add($FileCreationInfo)
                    $Ctx.Load($UploadFile)
                    $Ctx.ExecuteQuery()
                }

                #Compare the Attachment count in Item with Files in Folder
                If ($FilesOnItem -eq $FilesInFolder)
                {
                    $errorItemId += $Item.id
                }

                $DestFolder = $SharePointURL + "/" + $List.Title + "/" + $Item.id

                If ($Item.ID -like "IR*") {
                    $ResolutionDescription = ""
                    If ($Item.ResolutionDescription -ne $null -and $Item.ResolutionDescription.EndsWith("Attachments: "+ $DestFolder)){
                            $ResolutionDescription = $Item.ResolutionDescription
                    }
                    Else {
                        $ResolutionDescription = $Item.ResolutionDescription + " `r`n `r`nAttachments: "+ $DestFolder
                    }

                    $Item | Set-SCSMObject -Property ResolutionDescription -Value $ResolutionDescription
                }
                ElseIf ($Item.ID -like "CR*") {
                    $PostImplementationReview = ""
                    If ($Item.PostImplementationReview -ne $null -and $Item.PostImplementationReview.EndsWith("Attachments: "+ $DestFolder)){
                        $PostImplementationReview = $Item.PostImplementationReview
                    }
                    Else {
                        $PostImplementationReview = $Item.PostImplementationReview + " `r`n `r`nAttachments: "+ $DestFolder
                    }
        
                    $Item | Set-SCSMObject -Property PostImplementationReview -Value $PostImplementationReview               
                }
                ElseIf ($Item.ID -like "PR*") {
                    $ResolutionDescription = ""
                    If ($Item.ResolutionDescription -ne $null -and $Item.ResolutionDescription.EndsWith("Attachments: "+ $DestFolder)){
                        $ResolutionDescription = $Item.ResolutionDescription
                    }
                    Else {
                        $ResolutionDescription = $Item.ResolutionDescription + " `r`n `r`nAttachments: "+ $DestFolder
                    }
        
                    $Item | Set-SCSMObject -Property ResolutionDescription -Value $ResolutionDescription
                }
                ElseIf ($Item.ID -like "RR*") {
                    $PostImplementationReview = ""
                    If ($Item.PostImplementationReview -ne $null -and $Item.PostImplementationReview.EndsWith("Attachments: "+ $DestFolder)){
                        $PostImplementationReview = $Item.PostImplementationReview
                    }
                    Else {
                        $PostImplementationReview = $Item.PostImplementationReview + " `r`n `r`nAttachments: "+ $DestFolder
                    }
        
                    $Item | Set-SCSMObject -Property PostImplementationReview -Value $PostImplementationReview
                }
                ElseIf ($Item.ID -like "SR*") {
                    $ImplementationNotes = ""
                    If ($Item.Notes -ne $null -and $Item.Notes.EndsWith("Attachments: "+ $DestFolder)){
                        $ImplementationNotes = $Item.Notes
                    }
                    Else {
                        $ImplementationNotes = $Item.Notes + " `r`n `r`nAttachments: "+ $DestFolder
                    }
        
                    $Item | Set-SCSMObject -Property Notes -Value $ImplementationNotes
                }
            }
        } catch {
            $_ | Out-File -FilePath 'E:\SCAttachments\Error\error.txt' -Append
            Write-Host $List
        }
    }
}
# Output  
Get-ChildItem -Path $TempDir -Recurse| Foreach-object {Remove-item -Recurse -path $_.FullName }

$Output 
 
Remove-Module smlets
