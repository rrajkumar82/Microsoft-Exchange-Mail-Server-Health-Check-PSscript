# Using the FilterXML parameter:
$xmlIPAddress = @'
<QueryList>
   <Query Id="0" Path="Microsoft-IIS-Logging/Logs">
    <Select Path="Microsoft-IIS-Logging/Logs">
    *[System[(Level=1  or Level=2 or Level=3 or Level=4 or Level=0 or Level=5)]]
     and *[EventData[Data='/ecp/RulesEditor/TransportRules.slab' or Data='/ecp/RulesEditor/TransportRules.svc/GetList' or Data='/ecp/RulesEditor/ViewTransportRuleDetails.aspx' or Data='https://mail.domain.com/ecp/RulesEditor/TransportRules.slab?showhelp=false']]
     </Select>
  </Query>
</QueryList>
'@
$GetEvntTime = Get-WinEvent -FilterXml $xmlIPAddress -MaxEvents 1
#$Message = $Events.Message
$EventID = $GetEvntTime.Id
$MachineName = $GetEvntTime.MachineName
$RecordId = $GetEvntTime.RecordId
#$Source = $Events.ProviderName
$TimeCreated = $GetEvntTime.TimeCreated

#***************************************************************************************************************************
#Define the Date and Time
$timer = (Get-Date).tostring("dd-MM-yyyy,hhmmss")
# Define the report name
$reportname= "_" + $timer + "_RPT.csv";
$reportpath = "C:\Scripts\Event_Alert\Event_AlertReport\"
# Define the report path & name together
$ECPReport = $reportpath + $reportname;

#*************************************************************************************************************************** 
$Events = Get-WinEvent -FilterXml $xmlIPAddress -MaxEvents 15 | Select-Object *
# Parse out the event message data            
ForEach ($Event in $Events) 
{            
    # Convert the event to XML            
    $eventXML = [xml]$Event.ToXml()            
    # Iterate through each one of the XML message properties            
    For ($i=0; $i -lt $eventXML.Event.EventData.Data.Count; $i++) 
    {            
        # Append these as object properties            
        Add-Member -InputObject $Event -MemberType NoteProperty -Force ` 
            -Name  $eventXML.Event.EventData.Data[$i].name ` 
            -Value $eventXML.Event.EventData.Data[$i].'#text'            
    }            
}            

#****************************************************************************************************************************

$Events | Select-Object TimeCreated, RecordID, Message | Export-Csv -Path "$ECPReport" -NoTypeInformation

#****************************************************************************************************************************

$Subject ="E-MailFlow Rule | Alert From $MachineName"
$Body = "EventID: $EventID,`nRecordId: $RecordId,`nSource: $Source`nMachineName: $MachineName,`TimeCreated: $TimeCreated "

Send-MailMessage -From noreply@.com -Subject $Subject -To Ku@t.com -Body "<p> Hi, </p> <p> Some one accessing the email flow configuration webpage, please review the below details and find the attached report to get the Originating IP address details.. </p>`n$Body" -BodyAsHtml -Attachments $ECPReport  -SmtpServer mail.domain.com
    