# Created by : Ankur Ganveer
# Project : Data Repository
# Brief Description : fsac_placeholder script creates predefined folders & subfolders for required placeholder. 
#                     It also assign required permissions for A & B user AD groups on required placholder folder.


######### Pre-Requisite #########
# Compound folder must exists for which placeholder folder needs to be created.
# W: mapped on the server. This is the drive where folders & subfolders will be created.
# User account that will be use for execution of the script should not be local admin or remote access of the server.
# User account that will be use for execution of the script should have valid MS Office license.


######### Description #########
# fsac_placeholder script creates placeholder folder under existing compound folder. Existing folder name for compound and name for new placeholder folder is provided 
# during runtime of the script. The script also assign access permissions for A user AD group assign to the placeholder folder created.
# It also assign permission to the B AD user group to the study folder created.
# These folders are created in W: shared drive which is mapped on the server (X). 


######### Parameters #########
# Enter Compound name : (this is required compound folder name to be provided during runtime of the script)
# Enter placeholder name : (this is required study folder name to be provided during runtime of the script)
# Enter A user group name : (this is A group name who will need access to compound folder created during execution of the script)
# Enter B user group name : (this is B user group name who will need access to compound folder created during execution of the script)


# Release COM object. release excel objects

function Release-Ref ($ref) {

([System.Runtime.InteropServices.Marshal]::ReleaseComObject(

[System.__ComObject]$ref) -gt 0)

[System.GC]::Collect()

[System.GC]::WaitForPendingFinalizers()

}

# User input for Compound name, palceholder name and A user group 
# Only provided A user group and B user group will have access to given compound and study folders.

$Compound = Read-Host -Prompt 'Enter Compound Name:'

$Placeholder = Read-Host -Prompt 'Enter Placeholder Name:' 

$PrimaryUserG = Read-Host -Prompt 'A User Group:'


$Principal = "\$PrimaryUserG"

# this creates the directory

New-Item -Path "W:\A\$Compound\publications\$Placeholder" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\publications\$Placeholder\data" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\publications\$Placeholder\deliverables" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\publications\$Placeholder\programs" -Type Directory >$null 
New-Item -Path "W:\A\$Compound\publications\$Placeholder\outputs" -Type Directory >$null 

# Using CACLS to assign read access for A user group on particular folder locations.

ICACLS "W:\A\$Compound\publications\$Placeholder\data" /grant "${Principal}:(OI)(CI)(W,M)" >$null
ICACLS "W:\A\$Compound\publications\$Placeholder\deliverables" /grant "${Principal}:(OI)(CI)(W,M)" >$null
ICACLS "W:\A\$Compound\publications\$Placeholder\programs" /grant "${Principal}:(OI)(CI)(W,M)" >$null
ICACLS "W:\A\$Compound\publications\$Placeholder\outputs" /grant "${Principal}:(OI)(CI)(W,M)" >$null



 
