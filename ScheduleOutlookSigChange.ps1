# Original script from Reddit user u/andrew181082
# Modified by Ted & Austin (mostly Ted)
# Signature Creation 
#
#   Can be applied with GPO. If used in hosted environment, must be done from local server (where they log in, not from domain controller)
#   GPEDIT.MSC > Local Computer Policy > User Configuration > Windows Settings > Scriptes (Logon/Logoff) > Logon > Powershell Scripts
#
# NEW SIGNATURE
#
# **EITHER**     Note: your defualt signature folder is %AppData%\Microsoft\Signatures
#     Copy signature files: <new>.TXT, <new>.RTF & <new>.RTM files and the <new>_files directory
#       from: C:\Users\%UserYouAreCopyingFrom%\AppData\Roaming\Microsoft\Signatures to your Outlook Signatures folder
# **OR**
#     Copy an esisting signature TEMPLATE (one you've previously used) to your Outlook Signatures folder
#
# **EDIT Signature** Open Outlook > NewMessage > Signatures > Signatures
# **as needed** (depending on which copy step used above)
#     Rename the copied signature to <new>
#     Edit/Copy text and image(s). Add hyperlink(s). (DO NOT COPY IMAGES FROM A TICKETING SYSTEM - open the email link in Outlook and copy from there)
#     Change name to $DispalyName and title to $Title
#        (SELECT ONE LINE AT A TIME AND DO NOT SELECT THE SPACE AT END OF LINE - DOING SO MAY CHANGE PARAGRAPH/FONT STYLE)
# **SAVE**
# 
# **UPLOAD** the <new> files and <new>_files directory from your signature folder back to this folder 
#
# **EDIT Schedule** in this file
#     Copy last line and paste to end
#     remove comma comment from old last line and replace with a comma
#     update new last line with <new> signature name and go live date
# **SAVE edits** all done.


#
# EDIT THIS SECTION
#
$Schedule = (
#
# Make sure one of the dates is either today or before today!
# Go Live Dates must be arranged in ascending order
#
#  Go Live Date,    Signature,             version suffix (to force additional updates on same)
#("April 14, 2020", "Old Signature Name",   "4"),
#("April 22, 2020", "Another Old Name",     "1"),
#("April 28, 2020", "Not Needed",           "1"),
#("May 6, 2020",    "Filler",               "4"),
#("July 21, 2020",  "Foo",                  "2"),
#("Aug 10, 2020",   "Bar",                  "1"),
#("Aug 17, 2020",   "Stuff",                "1"),
#("Aug 24, 2020",   "Crap",                 "1"),
#("Aug 31, 2020",   "More crap",            "1"),
#("Sep  7, 2020",   "Blah",                 "1"),
#("Sep 18, 2020",   "Another Blah",         "1"),
("Sep 18, 2020",    "This Is Live",         "3"),
("Oct  9, 2025",    "This Is What Will be live on the day to the left", "1") # no comma after last row
)
$Reply = "Replies"
$ExemptUsers = @( "UsernameHere" ) # Remove Lockdown, don't lockdown & don't delete all signatures. Must be an array, empty @() or one element is OK
$TemplatesPath = "\\FileServer\AccessibleShare\EmailSignatures"
$SignaturesDir = "CustomSig"
#
# END OF EDIT SECTION
#

$today = get-date # "May 31, 2020" # test a different "today"
if ( $Schedule[0].count -eq 1 ) { # Fix PS issue: 2-dimensional array with only one row get inititalised as a 1 dim []
    $Schedule=(,$Schedule) } # cannot init above with "(," because it works with a single row but not more "(,(...),(...))"
foreach ($s in $Schedule) {
    if ( $today -gt (Get-Date $s[0]) ) {
        $Signature = $s[1]
        $VersionString = Get-Date $s[0] -Format ("yyyy.MM.dd." + $s[2]) }
    else {break}  }
$RootSignaturesPath = "${env:appdata}\Microsoft"
$UserName = $env:username
$BuildSignatures = @($Signature, $Reply)

$Extentions = @( ".txt", ".rtf", ".htm" )
$SignaturesPath = "$RootSignaturesPath\$SignaturesDir"
$VersionDir = "SigVersion"
$VersionFile = "$SignaturesPath\$VersionDir\version.txt"
$OfficeVersion = (Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer)."(Default)".split('.')[2] + ".0"
$RootRegKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Common\"
$RegKeyProperties = @(  
    ( "General",      "Signatures",     "$SignaturesDir", "string", $true  ),
    ( "MailSettings", "NewSignature",   "$Signature",     "string", $false ),
    ( "MailSettings", "ReplySignature", "$Reply",         "string", $false )
) | ForEach-Object {[pscustomobject]@{Path = $RootRegKey + $_[0]; Name = $_[1]; Value = $_[2]; Type = $_[3]; All = $_[4]}}

$Exempt = ($ExemptUsers -contains $UserName) -or # Domain Admin
    (((whoami /ALL /FO CSV | ConvertFrom-CSV) | Select 'User Name' | Where 'User Name' -like '*Domain Admins*') -ne $null)
# Remove Signature Lock Down for Exempt user
if ( $Exempt ) {
    foreach ($keyProp in $RegKeyProperties) { 
        if ( !$keyProp.All -and (Get-ItemProperty $keyProp.Path).PSObject.Properties.Name -contains $keyProp.Name ) {
            Remove-ItemProperty -Path $keyProp.Path -Name $keyProp.Name } } }

if ( (Test-Path -Path $VersionFile) -and (Get-Content $VersionFile) -eq $VersionString ) {
    Exit }

# ADUser values (via "AD-SI-" -> Active Directory - Service Interfaces - ...) instantiated with .filter
#   search: .Net class system.directoryservices.directorysearcher
$ADUser = ([adsisearcher]"(&(objectCategory=User)(samAccountName=$UserName))").FindOne().GetDirectoryEntry()
$DisplayName = $ADUser.givenName + $ADUser.sn
$Title = $ADUser.Title

if (Test-Path -Path $SignaturesPath) {
    if ( !$Exempt ) {
        Remove-Item $SignaturesPath\* -Recurse -Force } }
else {
    New-Item -Path $RootSignaturesPath -Name $SignaturesDir -ItemType Directory }

foreach ($sig in $BuildSignatures) {
    if (Test-Path -Path "$TemplatesPath\${sig}_files") {
        Copy-Item -Path "$TemplatesPath\${sig}_files" -Destination $SignaturesPath -Recurse -Force }
    foreach ($ext in $Extentions) {
        Invoke-Expression ('$SigOut = @"' + "`n" + (Get-Content -Path "$TemplatesPath\$sig$ext" | ForEach-Object { $_ + "`n" }) + "`n" + '"@')
        $SigOut | Out-File "$SignaturesPath\$sig$ext" -Force -Confirm:$false } }

foreach ($keyProp in $RegKeyProperties) {
    if ( $keyProp.All -or !$Exempt ) {
        New-ItemProperty -Path $keyProp.Path -Name $keyProp.Name -Value $keyProp.Value -PropertyType $keyProp.Type -Force } }

if ( !(Test-Path -Path "$SignaturesPath\$VersionDir") ) {
    New-Item -Path "$SignaturesPath" -Name $VersionDir -ItemType Directory }
$VersionString > $VersionFile
