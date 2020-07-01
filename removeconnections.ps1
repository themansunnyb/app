

#Choose pbix funtion
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PBIX (*.pbix)| *.pbix"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
#Error check function
function IsFileLocked([string]$filePath){
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}


#Choose file
try {$pathn = Get-FileName}
catch { "Incompatible File" }


#Check for errors
If([string]::IsNullOrEmpty($pathn )){            
    exit } 

elseif ( IsFileLocked($pathn) ){
    exit } 

#Run Script
else{    
   
    #Unzip pbix
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')
    $zipfile = $pathn.Substring(0,$pathn.Length-4) + "zip"
    Rename-Item -Path $pathn -NewName  $zipfile

    #Delete files
    $files   = 'Connections', 'SecurityBindings', 'DataModel'
    $stream = New-Object IO.FileStream($zipfile, [IO.FileMode]::Open)
    $mode   = [IO.Compression.ZipArchiveMode]::Update
    $zip    = New-Object IO.Compression.ZipArchive($stream, $mode)
    ($zip.Entries | ? { $files -contains $_.Name }) | % { $_.Delete() }

    #Close zip
    $zip.Dispose()
    $stream.Close()
    $stream.Dispose()

    #Repackage and open
    Rename-Item -Path $zipfile -NewName $pathn 
    Invoke-Item $pathn 
}