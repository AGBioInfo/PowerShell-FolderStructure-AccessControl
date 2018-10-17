######### Information #########
# FileName : fsac_compound.ps1
# Created by : Ankur Ganveer
# Project : Data Repository
# Brief Description : fsac_compound script creates predefined folders & subfolders for required compound & study. 
#                     It also assign required permissions for A & B user AD groups on required compound & study folders.


######### Pre-Requisite #########
# W: mapped on the server. This is the drive where folders & subfolders will be created.
# User account that will be use for execution of the script should not be local admin or remote access of the server.
# user account that will be use for execution of the script should have valid MS Office license.


######### Description #########
# fsac_compound script parse through the FolderStructure excel sheet and creates folder structure as defined in the excel sheet 
# for the compound name and study name provided as input during runtime of the script. The script also assign access permissions
# for A user AD group assign to the study folder created. It also assign permission to the B AD user group to the study folder created.
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

# User input for Compound name, study name, A user group and B user group
# Only provided A user group and B user group will have access to given compound and study folders.
 
$ParentTrue = Read-Host -Prompt 'Enter Compound Name:'

$StudyTrue = Read-Host -Prompt 'Enter Study Name:' 

$PrimaryUserG = Read-Host -Prompt 'A User Group:'

$BUser = Read-Host -Prompt 'B User Group :'

# Creating excel object

$objExcel = new-object -comobject excel.application 

$objExcel.Visible = $True 

 # Directory location where we have our excel files

$ExcelFilesLocation = “D:\FSAC Script\”

 # Open our excel file

$UserWorkBook = $objExcel.Workbooks.Open($ExcelFilesLocation + “FolderStructure3.xlsx”) 

# Item(1) refers to sheet 1 of of the workbook. 

$UserWorksheet = $UserWorkBook.Worksheets.Item(1)

# This is counter which will help to iterrate trough the loop. This is simply row count

# I am taking row count as 2, because the first row in my case is header. So we dont need to read the header data

$intRow = 2

Do {

 # Reading the first column of the current row. This is parent folder - compound folder. 

 $Parent = $UserWorksheet.Cells.Item($intRow, 1).Value()

 # Assiging compound name provided by user to compoundname1 

 If ($Parent -eq 'compoundname1') { $Parent = $ParentTrue }
 
 # Reading the second & third column of the current row. These values are for sub-folders. 

 $Child = $UserWorksheet.Cells.Item($intRow, 2).Value()
 $Child1 = $UserWorksheet.Cells.Item($intRow, 3).Value()

 # Assiging study name provided by user to Study1. This is study folder. 

 If ($Child1 -eq 'Study1') { $Child1 = $StudyTrue }

 # Reading the fourth & fifth column of the current row. These values are for sub-folders. 

 $Child2 = $UserWorksheet.Cells.Item($intRow, 4).Value()
 $Child3 = $UserWorksheet.Cells.Item($intRow, 5).Value()

 # this creates the directory 
 
 New-Item -Path "W:\A\$Parent\$Child\$Child1\$Child2\$Child3" -Type Directory >$null 

 # Move to next row

 $intRow++

 } While ($UserWorksheet.Cells.Item($intRow,1).Value() -ne $null)

 
# Exiting the excel object

$objExcel.Quit()

#Release all the objects used above

$a = Release-Ref($UserWorksheet)

$a = Release-Ref($UserWorkBook) 

$a = Release-Ref($objExcel)

# A user  AD group 

$Principal = "\$PrimaryUserG"

# B user  AD group

$B "\$BUser" 

# Using CACLS to assign read access for A & B user group on particular folder locations. 

ICACLS "W:\A\$ParentTrue\studies\$StudyTrue\final\pp_for_bdo" /grant "${B}:(OI)(CI)(RX)" >$null
ICACLS "W:\A\$ParentTrue\studies\$StudyTrue\preliminary\pp_for_bdo" /grant "${B}:(OI)(CI)(RX)" >$null
ICACLS "W:\A\$ParentTrue" /grant "${Principal}:(OI)(CI)(RX)" >$null
ICACLS "W:\A\$ParentTrue" /grant "${B}:(OI)(CI)(RX)" >$null

# Assigning write access for A user group on particular folder locations.

$StartingDir1 = "W:\A\$ParentTrue\studies\$StudyTrue\preliminary"
$StartingDir2 = "W:\A\$ParentTrue\studies\$StudyTrue\final"
$StartingDir3 = "W:\A\$ParentTrue\exploratoryanalysis"
$StartingDir4 = "W:\A\$ParentTrue\submissions\americas"
$StartingDir6 = "W:\A\$ParentTrue\submissions\asia pacific"
$StartingDir7 = "W:\A\$ParentTrue\submissions\emeaa"
$StartingDir8 = "W:\A\$ParentTrue\submissions\row"


foreach ($file in $(Get-ChildItem $StartingDir1 -recurse)) {
  
  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL
  
 }
 foreach ($file in $(Get-ChildItem $StartingDir2 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }

 foreach ($file in $(Get-ChildItem $StartingDir3 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }


 foreach ($file in $(Get-ChildItem $StartingDir4 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }

 
 foreach ($file in $(Get-ChildItem $StartingDir6 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }
 foreach ($file in $(Get-ChildItem $StartingDir7 -recurse)) {

  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }
 foreach ($file in $(Get-ChildItem $StartingDir8 -recurse)) {
  
  #ADD new permission with CACLS
  ICACLS $file.FullName /grant "${Principal}:(OI)(CI)(W,M)" >$NULL

 }


