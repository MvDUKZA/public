#. .\New-IsoFile.ps1
# New-IsoFile -SourcePath "C:\temp\vmtools_files" -OutputIso "C:\temp\vmtools.iso" -VolumeName "VMwareTools"

function New-IsoFile {
  param (
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$OutputIso,
    [string]$VolumeName = "VMwareTools"
  )
  $iso = New-Object -ComObject IMAPI2FS.MsftFileSystemImage
  $iso.VolumeName = $VolumeName
  $iso.FileSystemsToCreate = 3  # ISO9660 + Joliet
  $iso.AddTree($SourcePath, "/")
  $isoImage = $iso.CreateResultImage()
  $isoImage.ImageStream.CopyTo((New-Object -ComObject ADODB.Stream).Open(), 0)
  $isoImage.ImageStream.SaveToFile($OutputIso, 2)
}
