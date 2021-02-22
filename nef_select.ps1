
Function Select-FolderDialog
{
    param([string]$Description="Select Cherry Pick SRC",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

Function Select-FolderDialog1
{
    param([string]$Description="Select RAW SRC",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

Function Select-FolderDialog2
{
    param([string]$Description="Select RAW Cherry Destination",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}


Function Get-FileMetaData 
{ 
 <# 
  .Synopsis 
   This function gets file metadata and returns it as a custom PS Object  
  .Description 
   This function gets file metadata using the Shell.Application object and 
   returns a custom PSObject object that can be sorted, filtered or otherwise 
   manipulated. 
  .Example 
   Get-FileMetaData -folder "e:\music" 
   Gets file metadata for all files in the e:\music directory 
  .Example 
   Get-FileMetaData -folder (gci e:\music -Recurse -Directory).FullName 
   This example uses the Get-ChildItem cmdlet to do a recursive lookup of  
   all directories in the e:\music folder and then it goes through and gets 
   all of the file metada for all the files in the directories and in the  
   subdirectories.   
  .Example 
   Get-FileMetaData -folder "c:\fso","E:\music\Big Boi" 
   Gets file metadata from files in both the c:\fso directory and the 
   e:\music\big boi directory. 
  .Example 
   $meta = Get-FileMetaData -folder "E:\music" 
   This example gets file metadata from all files in the root of the 
   e:\music directory and stores the returned custom objects in a $meta  
   variable for later processing and manipulation. 
  .Parameter Folder 
   The folder that is parsed for files  
  .Notes 
   NAME:  Get-FileMetaData 
   AUTHOR: ed wilson, msft 
   LASTEDIT: 01/24/2014 14:08:24 
   KEYWORDS: Storage, Files, Metadata 
   HSG: HSG-2-5-14 
  .Link 
    Http://www.ScriptingGuys.com 
#Requires -Version 2.0 
#> 
Param([string[]]$folder) 
foreach($sFolder in $folder) 
 { 
  $a = 0 
  $objShell = New-Object -ComObject Shell.Application 
  $objFolder = $objShell.namespace($sFolder) 

  foreach ($File in $objFolder.items()) 
   {  
    $FileMetaData = New-Object PSOBJECT 
     for ($a ; $a  -le 266; $a++) 
      {  
        if($objFolder.getDetailsOf($File, $a)) 
          { 
            $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a))  = 
                  $($objFolder.getDetailsOf($File, $a)) } 
           $FileMetaData | Add-Member $hash 
           $hash.clear()  
          } #end if 
      } #end for  
    $a=0 
    $FileMetaData 
   } #end foreach $file 
 } #end foreach $sfolder 
} #end Get-FileMetaData

$jpg_src = Select-FolderDialog
$raw_src = Select-FolderDialog1 
$cherry_dest = Select-FolderDialog2
Write-host "Finding Rated JPGs..."
$ratedfiles = (Get-FileMetaData -folder $jpg_src | Where-Object {$_.Rating -eq "4 Stars"}) 

#$ratedfiles.Name
#Write-Host $cherry_src_fn
#Write-Host $raw_src

foreach ($file in $ratedfiles) {
    $name =""
    $basename = ""
    $name = $file.Name
    $basename = $name.Substring(0,$name.length-4)
    write-host "Copying" $raw_src"\"$basename".nef"
    Copy-Item $raw_src"\"$basename".nef" -Destination $cherry_dest
    #Write-Host $raw_src"\"$file".nef"
}

Write-Host "Copied "$ratedfiles.count" files."