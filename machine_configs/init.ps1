function draw-divider {
    param(
        [string] $title = ''
    )

    $side = "".PadLeft([Math]::Floor([decimal](((Get-Host).UI.RawUI.MaxWindowSize.Width - ($title | measure-object -character | select -expandproperty characters)) / 2)), "-")
    Write-Output ""
    Write-Output "$side$title$side"
    Write-Output ""
}

function refresh {
    RefreshEnv    
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User") 
}

draw-divider "getting user input for continuing unattended"
$firstName = Read-Host -Prompt 'First Name '
$email = Read-Host -Prompt 'Email '

draw-divider "installing chocolatey"

iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

refresh

draw-divider "installing chocolatey packages"

$chocoPackages = (
    "7zip.install",
    "adobereader",
    "awscli",
    "azure-functions-core-tools",
    "beyondcompare",
    "boostnote",
    #"ccleaner",
    "lastpass",
    "ditto",
    #"docker",
    "dotnet4.7",
    "dotnetcore-sdk",
    #"everything",
    "firefox",
    "git",
    "Git-Credential-Manager-for-Windows",
    "gitextensions",
    "golang",
    "GoogleChrome",
    "gpg4win",
    "habitat --version=0.79.1",
    "javaruntime",
    "jdk8",
    "jetbrainstoolbox",
    "jq",
    "microsoft-build-tools --version 14.0.25420.1", #2015
    "mRemoteNG",
    "netfx-4.7.1-devpack",
    "nodejs-lts --version=10.16.3",
    "NuGet.CommandLine",
    "NugetPackageExplorer",
    "openvpn",
    "packer",
    "poshgit",
    "postman",
    "putty",
    "python2",
    "python3",
    "redis-desktop-manager",
    "slack",
    #"spotify",
    "sql-server-management-studio",
    "SublimeText3",
    "terraform",
    "vagrant",
    "virtualbox",
    "VisualStudio2019Professional",
    #"visualstudio2019buildtools",
    "visualstudio2019-workload-data",
    "visualstudio2019-workload-manageddesktop",
    "visualstudio2019-workload-netcoretools",
    "visualstudio2019-workload-netcrossplat",
    "visualstudio2019-workload-netweb",
    "visualstudio2019-workload-node",
    "visualstudio2019-workload-python",
    "visualstudio2019-workload-netcorebuildtools",
    "visualstudio2019-workload-universal",
    "vscode",
    "vswhere",
    "webdeploy",
    "windirstat",
    "windows-sdk-10.1",
    "winscp",
    #"wox",
    "yarn",
    "yubico-authenticator"
)

# TODO: runs too slowly. should take more effort to configure installation as well
#    "visualstudio2015community",

$chocoPackages | ForEach-Object {
    draw-divider "choco -> $_"

    try {
        choco upgrade $_ -y
    }
    catch {
        draw-divider "Exception:"
        Write-Output $_.Exception | Format-List -force
    }
}

refresh

draw-divider "install vsts cred provider"

iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/Microsoft/artifacts-credprovider/master/helpers/installcredprovider.ps1'))

draw-divider "make chrome the default browser"

Add-Type -AssemblyName 'System.Windows.Forms'
Start-Process $env:windir\system32\control.exe -ArgumentList '/name Microsoft.DefaultPrograms /page pageDefaultProgram\pageAdvancedSettings?pszAppName=google%20chrome'
Sleep 2
[System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{TAB}{TAB}{TAB} ");

write-host "press any key to continue"
pause

draw-divider "installing vagrant plugins"

vagrant plugin install vagrant-reload

draw-divider "installing .net core tools"

dotnet tool install -g dotnet-outdated
dotnet tool install -g dotnet-serve
dotnet tool install -g dotnet-guid
dotnet tool install -g dotnet-script

draw-divider "installing vscode extensions"
#code --list-extensions

@(
    "bbenoist.vagrant",
    "bungcip.better-toml",
    "CoenraadS.bracket-pair-colorizer",
    "DotJoshJohnson.xml",
    "Gruntfuggly.todo-tree",
    "Gruntfuggly.vscode-journal-view",
    "humao.rest-client",
    "jmrog.vscode-nuget-package-manager",
    "k--kato.docomment",
    "mauve.terraform",
    "ms-azuretools.vscode-azureappservice",
    "ms-azuretools.vscode-azurefunctions",
    "ms-azuretools.vscode-azurestorage",
    "ms-azuretools.vscode-cosmosdb",
    "ms-mssql.mssql",
    "ms-vscode.azure-account",
    "ms-vscode.azurecli",
    "ms-vscode.csharp",
    "ms-vscode.go",
    "ms-vscode.PowerShell",
    "ms-python.python",
    "ms-vscode.vscode-node-azure-pack",
    "msazurermtools.azurerm-vscode-tools",
    "msjsdiag.debugger-for-chrome",
    "octref.vetur",
    "pajoma.vscode-journal",
    "PeterJausovec.vscode-docker",
    "robertohuertasm.vscode-icons",
    "sysoev.language-stylus",
    "tintoy.msbuild-project-tools",
    "WallabyJs.quokka-vscode"
) | ForEach-Object {
    draw-divider "vscode extension -> $_"

    try {
        code --install-extension $_
    }
    catch {
        draw-divider "Exception:"
        Write-Output $_.Exception | Format-List -force
    }
}

draw-divider "importing keys"

$mainKey = "0xBF6587060F4CBDBF"  #t

$keys = (
    $mainKey
)

$keys | ForEach-Object {
    draw-divider "importing key -> $_"

    try {
        gpg --recv $_
    }
    catch {
        draw-divider "Exception:"
        Write-Output $_.Exception | Format-List -force
    }
}

draw-divider "Configuring git"

git config --global user.name $firstName
git config --global user.email $email

git config --global commit.gpgsign true
git config --global user.signingkey $mainKey
git config --global gpg.program "C:/Program Files (x86)/GnuPG/bin/gpg.exe"

git config --global core.longpaths true

git config --global diff.tool bc
git config --global difftool.bc.path "C:/Program Files (x86)/Beyond Compare 4/bcomp.exe"

git config --global merge.tool bc
git config --global mergetool.bc.path "C:/Program Files (x86)/Beyond Compare 4/bcomp.exe"

draw-divider "Set Explorer prefs"

$key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
Set-ItemProperty $key Hidden 1
Set-ItemProperty $key HideFileExt 0
Stop-Process -processname explorer

draw-divider "Set application scaling prefs"
# https://superuser.com/a/1230356

$scalingFixTargets = (
    # "C:\Program Files\CCleaner\CCleaner.exe",
    # "C:\Program Files\CCleaner\CCleaner64.exe",
    "C:\Program Files (x86)\GitExtensions\GitExtensions.exe"
    # "C:\Program Files\Yubico\Yubico Authenticator\yubioath-desktop.exe" # doesn't do it... none of the options work well
)

$scalingFixLocation = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'

if (-not(Test-Path -Path $scalingFixLocation)) {
    New-Item -Path $scalingFixLocation
}

$scalingFixTargets | ForEach-Object {
    draw-divider "fixing scaling for -> $_"

    try {
        New-ItemProperty -Path $scalingFixLocation -Name $_ -PropertyType String -Value "~GDIDPISCALING DPIUNAWARE" -Force
    }
    catch {
        draw-divider "Exception:"
        Write-Output $_.Exception | Format-List -force
    }
}

draw-divider "enable windows features"

Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementConsole -All

draw-divider "installing misc MSIs"

$installers = (
    "https://s3.amazonaws.com/redshift-downloads/drivers/odbc/1.4.6.1000/AmazonRedshiftODBC32-1.4.6.1000.msi",
    "https://s3.amazonaws.com/redshift-downloads/drivers/odbc/1.4.6.1000/AmazonRedshiftODBC64-1.4.6.1000.msi",
    "https://download.microsoft.com/download/2/4/3/24374C5F-95A3-41D5-B1DF-34D98FF610A3/inetmgr_amd64_en-US.msi"
)

$installers | ForEach-Object {
    $outTemp = New-TemporaryFile | ForEach-Object { $_.FullName }
    $outTempMSI = [system.io.path]::ChangeExtension($outTemp, "msi")
    $outTempMSILog = [system.io.path]::ChangeExtension($outTemp, "log")

    Move-Item -Path $outTemp -Destination $outTempMSI

    try {
        Invoke-webrequest -uri $_ -OutFile $outTempMSI
        
        # https://stackoverflow.com/a/46224987/1301349
        # https://powershellexplained.com/2016-10-21-powershell-installing-msi-files/

        $MSIArguments = @(
            "/i"
            ('"{0}"' -f $outTempMSI)
            "/qn"
            "/norestart"
            "/L*v"
            $outTempMSILog
        )
        
        Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow 

        Write-Output "Installer log located in $outTempMSILog"
    }
    catch {
        Write-Output "Exception:"
        Write-Output $_.Exception | Format-List -force
    }
}

draw-divider "uninstalling windows apps"

$uninstallTargets = (
	"*3d*",
	"*alarm*",
	"*bing*",
	"*feedback*",
	"*help*",
	"*maps*",
	"*messaging*",
	"*music*",
	"*office*",
	"*people*",
	"*phone*",
	"*photos*",
	"*reality*",
	"*skypeapp*",
	"*solitaire*",
	"*speedtest*",
	"*sticky*",
	# "*store*",
	"*todos*",
	"*whiteboard*"
)

$ErrorActionPreference = 'SilentlyContinue'
$uninstallTargets | ForEach-Object {
    Get-AppxPackage $_ | Remove-AppxPackage | out-null
}

Write-Host "Uninstalling Teams"
# https://docs.microsoft.com/en-us/microsoftteams/scripts/powershell-script-teams-deployment-clean-up

$TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
$TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')

try
{
    if (Test-Path -Path $TeamsUpdateExePath) {
        Write-Host "Uninstalling Teams process"

        # Uninstall app
        $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
        $proc.WaitForExit()
    }
    if (Test-Path -Path $TeamsPath) {
        Write-Host "Deleting Teams directory"
        Remove-Item -Path $TeamsPath -Recurse
    }
}
catch
{
    Write-Error -ErrorRecord $_
}

Write-Host "Setting timezone..."
& "$env:windir\system32\tzutil.exe" /s "Eastern Standard Time"

Write-Host "Disabling Action Center..."
If (!(Test-Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer")) {
	New-Item -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" | Out-Null
}
Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "DisableNotificationCenter" -Type DWord -Value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications" -Name "ToastEnabled" -Type DWord -Value 0

draw-divider "yubikey sync"

write-host "enter yubikey then continue"
pause

gpg-connect-agent "scd serialno" "learn --force" /bye

draw-divider "installing additional TAP adapters"

. "C:\Program Files\TAP-Windows\bin\addtap.bat"
. "C:\Program Files\TAP-Windows\bin\addtap.bat"

draw-divider "configuring winrm"

Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private
winrm quickconfig
winrm set winrm/config/client '@{TrustedHosts="*"}'

draw-divider "install sublime context menu"

# from https://github.com/AtomHash/sublimetext3-contextmenu
$sublimeTempFile = "sublime.bat"

@'
@echo off
SET admin_st3_path=powershell cd 'c:\\program files\\sublime text 3\\'; Start-Process sublime_text.exe -Verb runAs
SET st3_path=C:\Program Files\Sublime Text 3\sublime_text.exe
SET st3_label_edit=Edit with Sublime Text
SET st3_label_admin_edit=Edit with Sublime Text as Admin
SET st3_label_admin_open=Open with Sublime Text as Admin
SET st3_label=Open with Sublime Text

rem add for all file types
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text"         /t REG_SZ /v "" /d "%st3_label_edit%" /f
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text\command" /t REG_SZ /v "" /d "%st3_path% \"%%1\"" /f
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text as Admin"         /t REG_SZ /v "" /d "%st3_label_admin_edit%" /f
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text as Admin"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text as Admin"         /t REG_SZ /v "Extended" /d ""
@reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Sublime Text as Admin\command" /t REG_SZ /v "" /d "%admin_st3_path% \"%%1\"" /f

rem add for folders
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open with Sublime Text"         /t REG_SZ /v "" /d "%st3_label%" /f
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open with Sublime Text"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open with Sublime Text\command" /t REG_SZ /v "" /d "%st3_path% \"%%V\"" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open with Sublime Text"         /t REG_SZ /v "" /d "%st3_label%" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open with Sublime Text"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open with Sublime Text\command" /t REG_SZ /v "" /d "%st3_path% \"%%V\"" /f

rem directory for admin
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open Sublime Text as Admin"         /t REG_SZ /v "" /d "%st3_label_admin_open%" /f
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open Sublime Text as Admin"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open Sublime Text as Admin"         /t REG_SZ /v "Extended" /d "" /f
@reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Open Sublime Text as Admin\command" /t REG_SZ /v "" /d "%admin_st3_path% \"%%V\"" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open Sublime Text as Admin"         /t REG_SZ /v "" /d "%st3_label_admin_open%" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open Sublime Text as Admin"         /t REG_EXPAND_SZ /v "Icon" /d "%st3_path%,0" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open Sublime Text as Admin"         /t REG_SZ /v "Extended" /d "" /f
@reg add "HKEY_CLASSES_ROOT\Directory\shell\Open Sublime Text as Admin\command" /t REG_SZ /v "" /d "%admin_st3_path% \"%%V\"" /f
'@ | out-file -encoding utf8 $sublimeTempFile

& ".\$sublimeTempFile"

Remove-Item ".\$sublimeTempFile"

draw-divider "go get latest driver for Quadro M2200"
write-host "as of 6/24: https://www.nvidia.com/download/driverResults.aspx/148432/en-us"
pause