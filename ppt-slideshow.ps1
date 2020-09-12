
$pptDest = "C:\temp\slideshow.ppsx"
$powerpoint = Get-Process powerpnt -ErrorAction SilentlyContinue


# Add the PowerPoint assemblies that we'll need
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

#Close PowerPoint if running
if ($powerpoint) {
    $powerpoint | Stop-Process -Force
    Start-Sleep -s 2
}

# Start PowerPoint
$ppt = new-object -com powerpoint.application
$ppt.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Set the locations where to find the PowerPoint files, and where to store the thumbnails
$pptPath = "\\someserver\someshare\powerpoints\"
$localPath = "C:\temp\"
$localFile = "slideshow"

# Loop through each PowerPoint File
Foreach($iFile in $(ls $pptPath -Filter "powerpoint.pptx")){

#Get-Member $iFile.FullName | Select-Object *
#Write-Host $ifile.Name

#Set-ItemProperty ($pptPath + $iFile) -name IsReadOnly -value $false
$filename = Split-Path $iFile -leaf
$file = $filename.Split(".")[0]
$oFile = $localPath + $localFile

# Open the PowerPoint file
$pres = $ppt.Presentations.Open($pptPath + $iFile)

# Now save it away as PDF 
$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLShow 
$pres.SaveAs($ofile,$opt)


# and Tidy-up 
$pres.Close();

}

#Clean Up
$ppt.quit();
$ppt = $null
[gc]::Collect();
[gc]::WaitForPendingFinalizers();


#Copy-Item $pptSource -Destination $pptDest -Force

Invoke-Item $pptDest
