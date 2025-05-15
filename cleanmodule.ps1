# Clean Az.Accounts from all known locations
Get-Module -Name Az.Accounts -ListAvailable |
    ForEach-Object { Remove-Item -Path $_.ModuleBase -Recurse -Force -ErrorAction SilentlyContinue }

# Clear NuGet package cache
Remove-Item "$env:LOCALAPPDATA\NuGet\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue

# Clean temp module folder if used
Remove-Item -Path "C:\Temp\PowerShellModules\Az.Accounts" -Recurse -Force -ErrorAction SilentlyContinue

# Reinstall cleanly
Install-Module Az.Accounts -Force -Scope CurrentUser -AllowClobber
