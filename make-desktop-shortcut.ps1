# Makes an "Offload" shortcut on your Desktop that opens the app in its own window (no browser, no console).
# Run it once: right-click this file > Run with PowerShell.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# Find pythonw (Python without a console window) through the "py" launcher.
$pythonw = & py -c "import sys,os; print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))"
if (-not (Test-Path $pythonw)) { Write-Host "Could not find Python (pythonw.exe). Install Python from python.org first."; exit 1 }

# An icon from the team's avatar, if Pillow is available.
$icon = Join-Path $here 'Offload.ico'
if (-not (Test-Path $icon)) {
    & py -c "from PIL import Image; im=Image.open(r'$here\avatars\manny.png').convert('RGBA'); s=max(im.size); c=Image.new('RGBA',(s,s),(0,0,0,0)); c.paste(im,((s-im.width)//2,(s-im.height)//2)); c.save(r'$icon', sizes=[(256,256),(64,64),(32,32),(16,16)])" 2>$null
}

$desktop = [Environment]::GetFolderPath('Desktop')
$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut((Join-Path $desktop 'Offload.lnk'))
$link.TargetPath = $pythonw
$link.Arguments = '"' + (Join-Path $here 'offload_desktop.pyw') + '"'
$link.WorkingDirectory = $here
$link.Description = 'Offload - your AI team'
if (Test-Path $icon) { $link.IconLocation = $icon }
$link.Save()
Write-Host "Done - there is now an Offload shortcut on your Desktop."
