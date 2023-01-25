# Instantiate a new Outlook object
$ol = new-object -comobject "Outlook.Application";

# Map to the MAPI namespace
$mapi = $ol.getnamespace("mapi");

$Items = $Mapi.Folders[1].Folders[33].Items

# Get a list of the Inbox folder items in the $Items variable
#$Items = $Mapi.GetDefaultFolder(4).Items;
#$folders = $Mapi.Folders[1].Folders | select folderpath;
