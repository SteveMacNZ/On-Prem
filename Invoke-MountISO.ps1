# make sure you adjust this path so it points to an
# existing ISO file:
$Path = "$env:temp\myImageFile.iso"
$result = Mount-DiskImage -ImagePath $Path -PassThru
$result

$volume = $result | Get-Volume
$letter = $volume.Driveletter + ":\"

explorer $letter 

# Dismount-DiskImage -ImagePath $Path 