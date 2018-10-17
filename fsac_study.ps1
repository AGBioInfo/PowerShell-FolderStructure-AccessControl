######### Information #########
# FileName : fsac_compound.ps1
# Created by : Ankur Ganveer
# Project : Data Repository 
# Brief Description : fsac_study script creates predefined folders & subfolders for required study. 
#                     It also assign required permissions for A & B user AD groups on required study folders.


######### Pre-Requisite #########
# Compound folder must exists for which study folder needs to be created.
# W: mapped on the server. This is the drive where folders & subfolders will be created.
# User account that will be use for execution of the script should not be local admin or remote access of the server.
# User account that will be use for execution of the script should have valid MS Office license.


######### Description #########
# fsac_study script creates study folder under existing compound folder. Existing folder name for compound and name for new study fodler is provided 
# during runtime of the script. The script also assign access permissions for A user AD group assign to the study folder created.
# It also assign permission to the B AD user group to the study folder created.
# These folders are created in W: shared drive which is mapped on the server (X). 


######### Parameters #########
# Enter Compound name : (this is required compound folder name to be provided during runtime of the script)
# Enter study name : (this is required study folder name to be provided during runtime of the script)
# Enter A user group name : (this is A group name who will need access to compound folder created during execution of the script)
# Enter B user group name : (this is B user group name who will need access to compound folder created during execution of the script)



# Release COM object. release excel objects

function Release-Ref ($ref) {

([System.Runtime.InteropServices.Marshal]::ReleaseComObject(

[System.__ComObject]$ref) -gt 0)

[System.GC]::Collect()

[System.GC]::WaitForPendingFinalizers()

}

# User input for existing compound name

$Compound = Read-Host -Prompt 'Enter Compound Name:'

# User input for new study folder that need to be created under existing compound folder 

$StudyUpdate = Read-Host -Prompt 'Enter Study Name:' 

# User input for A & B user group who will need access to new folder created

$PrimaryUserG = Read-Host -Prompt 'A User Group:'
$BUser = Read-Host -Prompt 'BUser Group :'

# A user  AD group 

$Principal = "\$PrimaryUserG"

# B user AD group

$B= "\$BUser"

# Follwoing creates new directories
 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\final\programs" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\final\outputs" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\final\data" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\final\deliverables" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\final\pp_for_bdo" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\preliminary\programs" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\preliminary\outputs" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\preliminary\data" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\preliminary\deliverables" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\studies\$StudyUpdate\preliminary\pp_for_bdo" -Type Directory >$null 

# Assigning read permissions for B user group

ICACLS "W:\A\$Compound\studies\$StudyUpdate\final\pp_for_bdo" /grant "${B}:(OI)(CI)(RX)" >$null
ICACLS "W:\A\$Compound\studies\$StudyUpdate\preliminary\pp_for_bdo" /grant "${B}:(OI)(CI)(RX)" >$null
ICACLS "W:\A\$ParentTrue" /grant "${B}:(OI)(CI)(RX)" >$null

# Assigning write permisisons to A user group

$StartingDir1 = "W:\A\$Compound\studies\$StudyUpdate\preliminary"
$StartingDir2 = "W:\A\$Compound\studies\$StudyUpdate\final"

foreach ($file in $(Get-ChildItem $StartingDir1 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }
 foreach ($file in $(Get-ChildItem $StartingDir2 -recurse)) {
  
  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }

 