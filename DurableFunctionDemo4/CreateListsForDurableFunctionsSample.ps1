[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
Connect-PNPOnline "https://tenant.sharepoint.com/sites/DurableFunctionDemo" -Credentials o365

New-PNPList -Title "Drafts" -Template DocumentLibrary -EnableVersioning -OnQuickLaunch
$draftlist = Get-PnPList "Drafts"
Add-PnPFieldFromXml -List $draftlist    '  <Field Type="User" 
  DisplayName="Document Owner" 
  Required="TRUE" 
  AddToDefaultView="TRUE"
  EnforceUniqueValues="FALSE" 
  StaticName="DocumentOwner" 
  Name="DocumentOwner" 
 />'
 Add-PnPFieldFromXml  -List $draftlist '  <Field Type="UserMulti" 
  DisplayName="StakeHolders" 
   Mult="TRUE" 
  Required="TRUE" 
    AddToDefaultView="TRUE"
  EnforceUniqueValues="FALSE" 
  StaticName="StakeHolders" 
  Name="StakeHolders" 
 />'

 Add-PnPField -DisplayName "Publish WF" -InternalName "workflowId" -Type Text  -List $draftlist -AddToDefaultView 


New-PNPList -Title "Published" -Template DocumentLibrary -EnableVersioning -OnQuickLaunch
$publisedList = Get-PnPList "Published"
Add-PnPFieldFromXml -List $publisedList    '  <Field Type="User" 
  DisplayName="Document Owner" 
  Required="TRUE" 
  AddToDefaultView="TRUE"
  EnforceUniqueValues="FALSE" 
  StaticName="DocumentOwner" 
  Name="DocumentOwner" 
 />'
 Add-PnPFieldFromXml  -List $publisedList '  <Field Type="UserMulti" 
  DisplayName="StakeHolders" 
   Mult="TRUE" 
  Required="TRUE" 
    AddToDefaultView="TRUE"
  EnforceUniqueValues="FALSE" 
  StaticName="StakeHolders" 
  Name="StakeHolders" 
 />'


New-PNPList -Title "Tasks" -Template GenericList -EnableVersioning -OnQuickLaunch
$taskslist = Get-PnPList "Tasks"
Add-PnPField -DisplayName "Workflow Id" -InternalName "workflowId" -Type Text  -List $taskslist -AddToDefaultView 
Add-PnPFieldFromXml -List $taskslist    '  <Field Type="User" 
  DisplayName="Assigned To" 
  Required="TRUE" 
  AddToDefaultView="TRUE"
  EnforceUniqueValues="FALSE" 
  StaticName="AssignedTo" 
  Name="AssignedTo" 
 />' 
 Add-PnPField -DisplayName "Action" -InternalName "Action" -Type Text  -List $taskslist -AddToDefaultView 
 Add-PnPField -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Approve","Reject"  -List $taskslist -AddToDefaultView 


New-PNPList -Title "Webhook" -Template GenericList -EnableVersioning -OnQuickLaunch

$webhook = Get-PnPList "Webhook"
Add-PnPListItem -List "Webhook" -Values @{"Title" = "will be set on first run"}




