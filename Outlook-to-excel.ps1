# Made by https://github.com/SuitableEmu
    <#
        .SYNOPSIS
        Reads from Outlook and parses the text from the body of the e-mail you want (some editing required)

        .DESCRIPTION
        Takes text from a specified e-mail (or multiple e-mails) and enters it into a text file.
}

#>


If(Get-InstalledModule ImportExcel) {
break
}
else{
Install-Module -Name ImportExcel
}

$Date = Get-Date -Format "dd.MM.yy"
$Year = Get-date -Format yyyy

#----------------------------------------- Functions ---------------------------------#

# This is what searches through your e-mail (it's really slow, it is set to use the default e-mail inbox)

Function Get-OutlookInBox

{

 Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null

 $olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]

 $outlook = new-object -comobject outlook.application

 $namespace = $outlook.GetNameSpace(“MAPI”)

 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

 $folder.items |


 Select-Object -Property Subject, ReceivedTime, Body

 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

} #end function Get-OutlookInbox


Function Clean-Memory {
Get-Variable |
 Where-Object { $startupVariables -notcontains $_.Name } |
 ForEach-Object {
  try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}
  catch { }
 }
}



#----------------------------------------- Outlook harvest ---------------------------------#


$inbox = Get-OutlookInbox

#----------------------------------------- Get-Info C450i ----------------------------------#

# Specify what it should look for in my case it's looking for the Subject "Counter List" and it has to have "c450i" within the body of the e-mail

$bodies = $inbox | Where-Object { $_.subject -match 'Counter List' -and $_.ReceivedTime -gt $Date -and $_.body -match "c450i" } | sort LastWriteTime | select -last 1 | Select-Object -ExpandProperty Body



#----------------------------------------- Parisng C450i ---------------------------------#

$bodies | Out-File C:\temp\tmpfile-C450i.txt

$tmpfile = "C:\Temp"
$file = Get-ChildItem $tmpfile | sort LastWriteTime | select -last 1 | Get-Content
$outfile01 = "C:\Temp\c450i.txt"

# if there are any special characters in the e-mail body you can exclude them here in my case it is only "," and it replaces it ( 'r' with a new line "n"

$file -replace ',',"`r`n" | Out-file $outfile01 -Force

#some cleanup

Remove-Item C:\temp\tmpfile-C450i.txt -Force

# this is where you set the tags and what content from the text file you want ( the first line is 0 )

$Model_Name = (Get-Content $outfile01)[0]
$Model_Tag = (Get-Content $outfile01)[1]
$Send_Date = $Date
$Total_Counter = (Get-Content $outfile01)[7]
$Total_Color_Counter = (Get-Content $outfile01)[9]
$Total_Black_Counter = (Get-Content $outfile01)[19]

# some trimming where i removed brackets from the text, this can also be added to filter out in $file -replace ...

$Model_Name = $Model_name.Trim('[]')

#---------------------------------------------------- c450i -------------------------------------------------------#
# Update the excel filename here

$filename = 'C:\temp\Namehere.xlsx'

$d = Get-Date -Format dd.MM

if($d -match 01.01){

$C450i=[pscustomobject]@{
"Model Name"=$Model_Tag;
"Send Date"=$Date;
"Total Counter Color"=$Total_Color_Counter;
"Total Counter Black"=$Total_Black_Counter;
"Total Counter"=$Total_Counter;

}

# formatting excel

$c450i | Export-excel $fileName -Append -Autosize -TableName A -TableStyle Medium16 -WorksheetName Worksheetnamegoeshere

}
elseif($d -match 01.02){

# Excel Name needs to go here

$Excel_Path = 'C:\temp\Namehere.xlsx'

$Excel = New-Object -Com Excel.Application
$Excel.Visible=$false

$Model_Name = (Get-Content C:\temp\C450i.txt)[0]
$Model_Name = $Model_name.Trim('[]')
$Model_Tag = (Get-Content C:\temp\C450i.txt)[1]
$Send_Date = $Date
$Total_Counter = (Get-Content C:\temp\C450i.txt)[7]
$Total_Color_Counter = (Get-Content C:\temp\C450i.txt)[9]
$Total_Black_Counter = (Get-Content C:\temp\C450i.txt)[19]


$Workbook = $Excel.Workbooks.Open($Excel_Path)
$page = "Name of sheet"+"-"+$year
$ws = $Workbook.worksheets | where-object {$_.Name -eq $page}
# Set variables for the worksheet cells, and for navigation

$cells=$ws.Cells
$cells.item(3,1) = Info
$cells.item(3,2) = Info
$cells.item(3,3) = Info
$cells.item(3,4) = Info
$cells.item(3,5) = Info
# Close the workbook and exit Excel
$workbook.Close($true)
$excel.quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null



}

else{

#error goes here

}


# general cleanup

Remove-Item C:\temp\c450i.txt


# Cleans memory

Clean-Memory
