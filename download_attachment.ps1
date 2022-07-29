<#
Bot Name    : Download attachments from email
Description : This script connects to respective user's mailbox in exchange server, filters the emails based on 
	          search criteria inputted and downloads the attachments from email.
Input       : username, username, Subject, Sender, FromDateTime, ToDateTime, Folder, DownloadPath, EWSdllpath, IsRead, FileExtension
Output      : Success/Failure
Version     : 1.0.1.20191205
#>
param(
[parameter(Mandatory=$true)] [string]$username,
[parameter(Mandatory=$true)] [string]$password,
[Parameter(Mandatory=$false)] [String] $mailsubject1 = '',
[Parameter(Mandatory=$false)] [String] $applenssubject = '',
[Parameter(Mandatory=$false)] [String] $IncidentID = '',
[Parameter(Mandatory=$false)] [String] $Subject = '',
[parameter(Mandatory=$false)] [string]$Sender = '',
[parameter (Mandatory=$false)] [string]$FromDateTime = '',
[parameter (Mandatory=$false)] [string]$ToDateTime = '',
[parameter (Mandatory=$false)] [string]$Folder = "Inbox",
[parameter (Mandatory=$false)] [string]$DownloadPath = "C:\",
[parameter (Mandatory=$true)] [string]$EWSdllpath,
[parameter (Mandatory=$false)] [string]$IsRead = "False",
[parameter (Mandatory=$false)] [string]$FileExtension
)

Function getFolderID {
param ([parameter(Mandatory)][Object]$service, [parameter(Mandatory)][string]$InpFolder)

try{
    $folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $folderview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
    $folderview.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)
    $folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep

    $folderResults = $service.FindFolders([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $folderview)
    foreach($folder in $folderResults.Folders){
        if ($folder.DisplayName -eq $InpFolder) {
            $folderID = $folder.ID
            return $folderID
         }
    }

    }
catch {
    $Response = "Error : Failed to get folder ID - " + $_.Exception.Message
    return $Response
}
}


try {
    # Load the Assemply
    [void][Reflection.Assembly]::LoadFile($EWSdllpath)


    # Create a new Exchange service object
    $service                     = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
    $Service.Credentials         = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($username,$password)
    $service.Url                 = "https://outlook.office365.com/EWS/Exchange.asmx"
    $search_filter    = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);

    if ($sender.Trim() -eq '' -and $Subject.Trim() -eq '' -and $FromDateTime.Trim() -eq '' -and $IsRead -eq 'True')
    {
        return "Please input any one search filter (Subject/Sender/FromDateTime)" }
    if ($FromDateTime.Trim() -eq '' -and $ToDateTime.Trim() -ne '')
    {
        return "Please input FromDateTime along with ToDateTime" }


    if ($sender) {
        $search_filter_for_sender     = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender, $sender)
        $search_filter.add($search_filter_for_sender)
        }
    if ($subject) {
        $search_filter_for_subject     = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, $subject)
        $search_filter.add($search_filter_for_subject)
        }

    if ($FromDateTime) {
        $FromDateTime = Get-Date -Date $FromDateTime -Format "MM/dd/yyyy h:mm:ss tt" -ErrorAction Stop

        $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
        $DateFormat = $CultureDateTimeFormat.ShortDatePattern
        $TimeFormat = $CultureDateTimeFormat.LongTimePattern
        $DateTimeFormat = "$DateFormat $TimeFormat"

        $FormatFromDateTime = [system.DateTime]::ParseExact($FromDateTime, $DateTimeFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)
        $search_filter_for_fromtime =  new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$FormatFromDateTime)
        $search_filter.add($search_filter_for_fromtime)
        }
    if ($ToDateTime) {
        $ToDateTime = Get-Date -Date $ToDateTime -Format "MM/dd/yyyy h:mm:ss tt" -ErrorAction Stop
        $FormatToDateTime = [system.DateTime]::ParseExact($ToDateTime, $DateTimeFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)

        $search_filter_for_totime =  new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,$FormatToDateTime)

        $search_filter.add($search_filter_for_totime)
        }
    if ($IsRead -eq "False") {
        $search_filter_for_unread_mails =  new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $IsRead)
        $search_filter.add($search_filter_for_unread_mails)
        }

    $search_filter_for_attachment = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, "True")
    $search_filter.add($search_filter_for_attachment)

    $FolderID = getFolderID -service $service -InpFolder $Folder

    # create Property Set to include body and header of email
    $pageSize = 1000
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize)
    $itemView.PropertySet =  New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    # set email body to text
    $itemView.PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
    $filterResults = $service.FindItems($FolderID, $search_filter, $itemView)
    if ($filterResults.TotalCount -eq 0) {
        return "No new emails matching the input criteria!!"
    }
	if($FileExtension) {
        if($FileExtension.StartsWith('*')) {
            return "Please enter the file extension in format (.txt, .docx, .xlsx, etc.,)" }
		else {
        $FileExtension += "," }
    }
	
	if(-not(Test-Path $DownloadPath)) { return "Please enter valid path to save the downloaded files!!" }

    $attachment_flag = 0
    $attach_count = 0
    # Do/while loop for paging through the folder
    do
    {
        foreach ($item in $filterResults.Items)
        {

            if ($item.HasAttachments -eq "True") {
                $count += 1
                #set attachment flag
                $attachment_flag = 1

                # Output the results
                $response += "Reading email - ($count)`n " +
                             "From : $($item.From.Name) `n " +
                             "Subject: $($item.Subject) `n " +
                             "Received at: $($item.DateTimeReceived) `n "

                # load the additional properties for the item
                $item.Load($itemView.propertySet)
                $sub = "Subject " + $item.Subject

               
                foreach($attach in $item.Attachments) {

                    #Download the attachements
                    $attach.Load()
                    if( -not $attach.IsInline) {
                        $attach_count += 1
                        $attachName = $attach.Name
                        
                        if($FileExtension) {
                            if($FileExtension -notmatch ([System.IO.Path]::GetExtension($attachName)).ToLower())
                            {
                                $response += "Attachment of type - " + [System.IO.Path]::GetExtension($attachName) + " is not requested for download!!"
                                $response += "`n"
                                continue
                            }                            
                        }
                        

                        $download_count += 1
                        if ($attach.GetType().name -eq "FileAttachment") {
                            $fiFile = new-object System.IO.FileStream(($DownloadPath + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
                                    $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                                    $fiFile.Close()
                                    $response += "Downloaded Attachment [" + ($download_count) + "] to : " + (($DownloadPath + '\' + $attach.Name.ToString()))
                            $response += "`n"
                        }
                        elseif ($attach.GetType().Name -eq "ItemAttachment") {
                            $mimePropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
                            $attach.Load($mimePropertySet)
                            $attachmentData = $attach.Item.MimeContent.Content
                            $attachname = $attach.Name.ToString() -replace '[\W]', ' '
                            $fiFile = new-object System.IO.FileStream(($DownloadPath + "\" + $attachname + ".eml"), [System.IO.FileMode]::Create)
                                $fiFile.Write($attachmentData, 0, $attachmentData.Length)
                            $fiFile.Close()
                            $response += "Downloaded Attachment [" + ($download_count) + "] to : " + (($DownloadPath + '\' + $attachname + ".msg"))
                            $response += "`n"
                         }
                         else {
                            $response += "Attachment - " + $attach.Name.ToString() + " is of " +  $attach.GetType().Name + ". It cannot be downloaded by this bot. "
                            $response += "`n"
                         }
                    }
                }
                if($IsRead -eq "False" -and ($download_count -ne 0)) {
                    $item.isRead = $true
                    $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve);     
                } 
            }

         }
    } while ($filterResults.MoreAvailable)
    if ($filterResults.Items.Count -eq 0) {
        $response = "No emails found for the input criteria!!"
        }
    if ($download_count -eq 0) {
        $response = "No attachments to download!!"
        }
    
    return $response
}
catch {
    $Response = "Error : " + $_.Exception.Message +$_.InvocationInfo.ScriptLineNumber
    return $Response
}
finally {
    if ($service -ne $null) {Remove-Variable service }
    if ($search_filter -ne $null) { Remove-Variable search_filter }
    if($FolderID -ne $null) { Remove-Variable FolderID }
    if ($response -ne $null) { Remove-Variable response }
}
