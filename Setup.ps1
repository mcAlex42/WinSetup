
# Elevation and globals -------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	if ($PSCommandPath) {
		Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
		exit
	}
	Write-Error "Administrator privileges are required."
	exit 1
}

$ErrorActionPreference = 'Stop'
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-ProgressLog {
	param([string]$Message,[string]$Level = 'INFO')
	$ts = (Get-Date).ToString('HH:mm:ss')
	Write-Host "[$ts][$Level] $Message"
}

# Load UI automation assemblies once
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

# Helpers --------------------------------------------------------------
function Get-UIAElement {
	param(
		[System.Windows.Automation.AutomationElement]$Root,
		[string]$AutomationId,
		[string]$Name,
		[int]$TimeoutSeconds = 15
	)
	$conditions = @()
	if ($AutomationId) {
		$conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty,$AutomationId)
	}
	if ($Name) {
		$conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty,$Name)
	}
	if (-not $conditions) { return $null }
	$condition = if ($conditions.Count -eq 1) { $conditions[0] } else { New-Object System.Windows.Automation.AndCondition($conditions) }
	$start = Get-Date
	do {
		$found = $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants,$condition)
		if ($found) { return $found }
		Start-Sleep -Milliseconds 300
	} while ((Get-Date) -lt $start.AddSeconds($TimeoutSeconds))
	return $null
}

function Invoke-UIAElement {
	param([System.Windows.Automation.AutomationElement]$Element)
	if (-not $Element) { return }
	try {
		$invoke = $null
		if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern,[ref]$invoke)) { $invoke.Invoke(); return }
	} catch {}
	try {
		$toggle = $null
		if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$toggle)) { $toggle.Toggle(); return }
	} catch {}
	try {
		$select = $null
		if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$select)) { $select.Select(); return }
	} catch {}
	try {
		$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
		$parent = $walker.GetParent($Element)
		if ($parent) {
			$selectParent = $null
			if ($parent.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$selectParent)) { $selectParent.Select(); return }
		}
	} catch {}
	try {
		if ($Element.Current.IsKeyboardFocusable) {
			$Element.SetFocus(); Start-Sleep -Milliseconds 100; [System.Windows.Forms.SendKeys]::SendWait("{ENTER}"); return
		}
	} catch {}
}

function Set-RegistryValueSafe {
	param([string]$Path,[string]$Name,[string]$Type='DWord',[Object]$Value)
	$psPath = if ($Path -like 'HKCU:*' -or $Path -like 'HKLM:*') { $Path } else { $Path.Replace('HKCU','HKCU:').Replace('HKLM','HKLM:') }
	if (-not (Test-Path $psPath)) { New-Item -Path $psPath -Force | Out-Null }
	try { New-ItemProperty -Path $psPath -Name $Name -PropertyType $Type -Value $Value -Force | Out-Null } catch { Write-ProgressLog "Registry write failed ${psPath}/${Name}: $_" 'WARN' }
}

function Invoke-DismSafe {
	param([string]$Arguments)
	try { Start-Process -FilePath dism.exe -ArgumentList $Arguments -NoNewWindow -Wait -PassThru | Out-Null } catch { Write-ProgressLog "DISM failed: $Arguments ($_ )" 'WARN' }
}

function Invoke-External {
	param([string]$FilePath,[string]$Arguments)
	try { Start-Process -FilePath $FilePath -ArgumentList $Arguments -NoNewWindow -Wait -PassThru | Out-Null } catch { Write-ProgressLog "Command failed: $FilePath $Arguments ($_ )" 'WARN' }
}

function Save-PTJson {
	param([string]$Path,[object]$JsonObject)
	$jsonString = $JsonObject | ConvertTo-Json -Depth 10
	$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
	[System.IO.File]::WriteAllText($Path,$jsonString,$utf8NoBom)
}

function Close-SettingsWindow {
	param([System.Windows.Automation.AutomationElement]$Window)
	if (-not $Window) { return }
	try {
		$pattern = $null
		if ($Window.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern,[ref]$pattern)) { $pattern.Close(); return }
	} catch {}
	try { $Window.SetFocus(); [System.Windows.Forms.SendKeys]::SendWait('%{F4}') } catch {}
	Stop-Process -Name 'SystemSettings' -Force -ErrorAction SilentlyContinue
}

function Invoke-MinimizeAllWindows { try { (New-Object -ComObject Shell.Application).MinimizeAll() | Out-Null } catch {} }

function Test-WingetAvailable {
	$wg = Get-Command winget -ErrorAction SilentlyContinue
	if (-not $wg) { throw "winget is not available" }
}

function Wait-JobWithOutput {
	param([System.Collections.ArrayList]$Jobs)
	while ($Jobs.Count -gt 0) {
		$done = $Jobs | Wait-Job -Any
		if ($done) {
			Receive-Job -Job $done -Keep | ForEach-Object { Write-ProgressLog $_ 'JOB' }
			$Jobs.Remove($done) | Out-Null
			Remove-Job $done -Force -ErrorAction SilentlyContinue
		}
	}
}

# Functions ------------------------------------------------------------
function Set-LockScreen {

	Write-ProgressLog "Setting lock screen"

    $lockScreenImage = "C:\Windows\Web\Screen\img102.jpg"
    if (-not (Test-Path $lockScreenImage)) {
        Write-Warning "Lock screen image not found: $lockScreenImage"
        return
    }

    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    $cdmPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    $cdmSubscriptionFlag = "SubscribedContent-338387Enabled"
    $creativePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Creative\$sid"
    $personalizationPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
    $personalizationPolicyUser = "HKCU:\Software\Policies\Microsoft\Windows\Personalization"
    $cloudPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    $personalizationCsp = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    $personalizationCspMachine = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

    New-Item -Path $cdmPath -Force | Out-Null
    New-Item -Path $creativePath -Force | Out-Null
    New-Item -Path $personalizationPolicy -Force | Out-Null
    New-Item -Path $personalizationPolicyUser -Force | Out-Null
    New-Item -Path $cloudPolicy -Force | Out-Null
    New-Item -Path $personalizationCsp -Force | Out-Null
    New-Item -Path $personalizationCspMachine -Force | Out-Null

    # Turn off Windows Spotlight and lock screen tips
    Set-RegistryValueSafe -Path $cdmPath -Name "RotatingLockScreenEnabled" -Value 0 -Type DWord
    Set-RegistryValueSafe -Path $cdmPath -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord
    Set-RegistryValueSafe -Path $cdmPath -Name $cdmSubscriptionFlag -Value 0 -Type DWord
    Set-RegistryValueSafe -Path $creativePath -Name "RotatingLockScreenEnabled" -Value 0 -Type DWord
    Set-RegistryValueSafe -Path $creativePath -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord

    # Enforce picture lock screen and supply the image path
    Set-RegistryValueSafe -Path $personalizationPolicy -Name "LockScreenImage" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationPolicyUser -Name "LockScreenImage" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationPolicy -Name "NoLockScreenSlideshow" -Value 1 -Type DWord
    Set-RegistryValueSafe -Path $personalizationPolicyUser -Name "NoLockScreenSlideshow" -Value 1 -Type DWord
    Set-RegistryValueSafe -Path $cloudPolicy -Name "DisableWindowsSpotlightOnLockScreen" -Value 1 -Type DWord
    Set-RegistryValueSafe -Path $cloudPolicy -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord

    # Mirror path in user-scoped personalization to nudge shell to reload
    Set-RegistryValueSafe -Path $personalizationCsp -Name "LockScreenImagePath" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationCsp -Name "LockScreenImageUrl" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationCsp -Name "LockScreenImageStatus" -Value 1 -Type DWord
    Set-RegistryValueSafe -Path $personalizationCsp -Name "LockScreenImageOptions" -Value 0 -Type DWord
    Set-RegistryValueSafe -Path $personalizationCspMachine -Name "LockScreenImagePath" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationCspMachine -Name "LockScreenImageUrl" -Value $lockScreenImage -Type String
    Set-RegistryValueSafe -Path $personalizationCspMachine -Name "LockScreenImageStatus" -Value 1 -Type DWord
    Set-RegistryValueSafe -Path $personalizationCspMachine -Name "LockScreenImageOptions" -Value 0 -Type DWord

    & rundll32.exe user32.dll,UpdatePerUserSystemParameters 1, True | Out-Null
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
}

function Set-WindowsTheme {
    Write-ProgressLog "Setting windows theme"

    # 1. Define Paths and Values
    $themeAPath = Join-Path $env:WINDIR "Resources\Themes\themeA.theme"
    $dwmPath = "HKCU:\Software\Microsoft\Windows\DWM"
    $accentPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent"
    $themeHistoryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\History"
    $personalizePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $themesRoot = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes"

    $accentColor = 4292114432
    $startColorMenu = 4290799360
    $colorizationColor = 3288365268
    $accentPalette = [byte[]](0x99,0xEB,0xFF,0x00,0x4C,0xC2,0xFF,0x00,0x00,0x91,0xF8,0x00,0x00,0x78,0xD4,0x00)

    # 2. Apply the Theme File First
    if (Test-Path $themeAPath) {
        Start-Process -FilePath $themeAPath -Wait
        Start-Sleep -Milliseconds 200
        
        Set-RegistryValueSafe -Path $themesRoot -Name "CurrentTheme" -Value $themeAPath -Type String
        Set-RegistryValueSafe -Path $themesRoot -Name "ThemeMRU" -Value "$themeAPath;" -Type String
    } else {
        Write-Warning "Theme file not found: $themeAPath"
    }

    # 3. Apply Custom Registry Tweaks (Overrides theme defaults)
    Set-RegistryValueSafe -Path $themeHistoryPath -Name "AutoColor" -Value 0 -Type DWord

    # DWM (Desktop Window Manager) tweaks
    Set-RegistryValueSafe -Path $dwmPath -Name "AccentColor" -Value $accentColor
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationColor" -Value $colorizationColor
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationColorBalance" -Value 89
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationAfterglow" -Value $colorizationColor
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationAfterglowBalance" -Value 10
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationBlurBalance" -Value 1
    Set-RegistryValueSafe -Path $dwmPath -Name "EnableWindowColorization" -Value 1
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorizationGlassAttribute" -Value 1
    Set-RegistryValueSafe -Path $dwmPath -Name "ColorPrevalence" -Value 0

    # Accent and Menu tweaks
    Set-RegistryValueSafe -Path $accentPath -Name "AccentColorMenu" -Value $accentColor
    Set-RegistryValueSafe -Path $accentPath -Name "StartColorMenu" -Value $startColorMenu
    Set-RegistryValueSafe -Path $accentPath -Name "AccentPalette" -Value $accentPalette -Type Binary

    # Dark Mode and Personalization
    Set-RegistryValueSafe -Path $personalizePath -Name "ColorPrevalence" -Value 0
    Set-RegistryValueSafe -Path $personalizePath -Name "AppsUseLightTheme" -Value 0
    Set-RegistryValueSafe -Path $personalizePath -Name "SystemUsesLightTheme" -Value 0

    Set-LockScreen

    # 4. Refresh Shell
    & rundll32.exe user32.dll,UpdatePerUserSystemParameters 1, True | Out-Null
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Stop-Process -Name 'SystemSettings' -Force -ErrorAction SilentlyContinue
}

function Set-PrivacySettings {
	Write-ProgressLog "Applying O&O ShutUp10++ config"
	$tempDir = $env:TEMP
	$ooExe = Join-Path $tempDir 'OOSU10.exe'
	$ooCfg = Join-Path $tempDir 'ooshutup10.cfg'
	$ooConfigContent = @'
	############################################################################
	# This file was created with O&O ShutUp10++ V2.1.1015
	# and can be imported onto another computer. 
	#
	# Download the application at https://www.oo-software.com/shutup10
	# You can then import the file from within the program. 
	#
	# Alternatively you can import it automatically over a command line.
	# Simply use the following parameter: 
	# OOSU10.exe <path to file>
	# 
	# Selecting the Option /quiet ends the app right after the import and the
	# user does not get any feedback about the import.
	#
	# We are always happy to answer any questions you may have!
	# © 2015-2025 O&O Software GmbH, Berlin. All rights reserved.
	# https://www.oo-software.com/
	############################################################################

	P001	+	# Disable sharing of handwriting data (Category: Privacy)
	P002	+	# Disable sharing of handwriting error reports (Category: Privacy)
	P003	+	# Disable Inventory Collector (Category: Privacy)
	P004	-	# Disable camera in logon screen (Category: Privacy)
	P005	+	# Disable and reset Advertising ID and info for the machine (Category: Privacy)
	P006	+	# Disable and reset Advertising ID and info for current user (Category: Privacy)
	P008	+	# Disable transmission of typing information (Category: Privacy)
	P026	+	# Disable advertisements via Bluetooth (Category: Privacy)
	P027	+	# Disable the Windows Customer Experience Improvement Program (Category: Privacy)
	P028	+	# Disable backup of text messages into the cloud (Category: Privacy)
	P064	+	# Disable suggestions in the timeline (Category: Privacy)
	P065	+	# Disable suggestions in Start (Category: Privacy)
	P066	+	# Disable tips, tricks, and suggestions when using Windows (Category: Privacy)
	P067	+	# Disable showing suggested content in the Settings app (Category: Privacy)
	P070	+	# Disable the possibility of suggesting to finish the setup of the device (Category: Privacy)
	P069	+	# Disable Windows Error Reporting (Category: Privacy)
	P009	-	# Disable biometrical features (Category: Privacy)
	P010	-	# Disable app notifications (Category: Privacy)
	P015	-	# Disable access to local language for browsers (Category: Privacy)
	P068	-	# Disable text suggestions when typing on the software keyboard (Category: Privacy)
	P016	-	# Disable sending URLs from apps to Windows Store (Category: Privacy)
	A001	+	# Disable recordings of user activity (Category: Activity History and Clipboard)
	A002	+	# Disable storing users' activity history on this device (Category: Activity History and Clipboard)
	A003	+	# Disable the submission of user activities to Microsoft (Category: Activity History and Clipboard)
	A004	-	# Disable storage of clipboard history for whole machine (Category: Activity History and Clipboard)
	A006	-	# Disable storage of clipboard history for current user (Category: Activity History and Clipboard)
	A005	-	# Disable the transfer of the clipboard to other devices via the cloud (Category: Activity History and Clipboard)
	P007	-	# Disable app access to user account information on this device (Category: App Privacy)
	P036	-	# Disable app access to user account information for current user (Category: App Privacy)
	P025	-	# Disable Windows tracking of app starts (Category: App Privacy)
	P033	-	# Disable app access to diagnostics information on this device (Category: App Privacy)
	P023	-	# Disable app access to diagnostics information for current user (Category: App Privacy)
	P056	-	# Disable app access to device location on this device (Category: App Privacy)
	P057	-	# Disable app access to device location for current user (Category: App Privacy)
	P012	-	# Disable app access to camera on this device (Category: App Privacy)
	P034	-	# Disable app access to camera for current user (Category: App Privacy)
	P013	-	# Disable app access to microphone on this device (Category: App Privacy)
	P035	-	# Disable app access to microphone for current user (Category: App Privacy)
	P062	-	# Disable app access to use voice activation for current user (Category: App Privacy)
	P063	-	# Disable app access to use voice activation when device is locked for current user (Category: App Privacy)
	P081	-	# Disable the standard app for the headset button (Category: App Privacy)
	P047	-	# Disable app access to notifications on this device (Category: App Privacy)
	P019	-	# Disable app access to notifications for current user (Category: App Privacy)
	P048	-	# Disable app access to motion on this device (Category: App Privacy)
	P049	-	# Disable app access to movements for current user (Category: App Privacy)
	P020	-	# Disable app access to contacts on this device (Category: App Privacy)
	P037	-	# Disable app access to contacts for current user (Category: App Privacy)
	P011	-	# Disable app access to calendar on this device (Category: App Privacy)
	P038	-	# Disable app access to calendar for current user (Category: App Privacy)
	P050	-	# Disable app access to phone calls on this device (Category: App Privacy)
	P051	-	# Disable app access to phone calls for current user (Category: App Privacy)
	P018	-	# Disable app access to call history on this device (Category: App Privacy)
	P039	-	# Disable app access to call history for current user (Category: App Privacy)
	P021	-	# Disable app access to email on this device (Category: App Privacy)
	P040	-	# Disable app access to email for current user (Category: App Privacy)
	P022	-	# Disable app access to tasks on this device (Category: App Privacy)
	P041	-	# Disable app access to tasks for current user (Category: App Privacy)
	P014	-	# Disable app access to messages on this device (Category: App Privacy)
	P042	-	# Disable app access to messages for current user (Category: App Privacy)
	P052	-	# Disable app access to radios on this device (Category: App Privacy)
	P053	-	# Disable app access to radios for current user (Category: App Privacy)
	P054	-	# Disable app access to unpaired devices on this device (Category: App Privacy)
	P055	-	# Disable app access to unpaired devices for current user (Category: App Privacy)
	P029	-	# Disable app access to documents on this device (Category: App Privacy)
	P043	-	# Disable app access to documents for current user (Category: App Privacy)
	P030	-	# Disable app access to images on this device (Category: App Privacy)
	P044	-	# Disable app access to images for current user (Category: App Privacy)
	P031	-	# Disable app access to videos on this device (Category: App Privacy)
	P045	-	# Disable app access to videos for current user (Category: App Privacy)
	P032	-	# Disable app access to the file system on this device (Category: App Privacy)
	P046	-	# Disable app access to the file system for current user (Category: App Privacy)
	P058	-	# Disable app access to wireless equipment on this device (Category: App Privacy)
	P059	-	# Disable app access to wireless technology for current user (Category: App Privacy)
	P060	-	# Disable app access to eye tracking on this device (Category: App Privacy)
	P061	-	# Disable app access to eye tracking for current user (Category: App Privacy)
	P071	-	# Disable the ability for apps to take screenshots on this device (Category: App Privacy)
	P072	-	# Disable the ability for apps to take screenshots for current user (Category: App Privacy)
	P073	-	# Disable the ability for desktop apps to take screenshots for current user (Category: App Privacy)
	P074	-	# Disable the ability for apps to take screenshots without borders on this device (Category: App Privacy)
	P075	-	# Disable the ability for apps to take screenshots without borders for current user (Category: App Privacy)
	P076	-	# Disable the ability for desktop apps to take screenshots without margins for current user (Category: App Privacy)
	P077	-	# Disable app access to music libraries on this device (Category: App Privacy)
	P078	-	# Disable app access to music libraries for current user (Category: App Privacy)
	P079	-	# Disable app access to downloads folder on this device (Category: App Privacy)
	P080	-	# Disable app access to downloads folder for current user (Category: App Privacy)
	P024	-	# Prohibit apps from running in the background (Category: App Privacy)
	S001	-	# Disable password reveal button (Category: Security)
	S002	-	# Disable user steps recorder (Category: Security)
	S003	+	# Disable telemetry (Category: Security)
	S008	-	# Disable Internet access of Windows Media Digital Rights Management (DRM) (Category: Security)
	E101	+	# Disable tracking in the web (Category: Microsoft Edge (new version based on Chromium))
	E201	-	# Disable tracking in the web (Category: Microsoft Edge (new version based on Chromium))
	E115	-	# Disable check for saved payment methods by sites (Category: Microsoft Edge (new version based on Chromium))
	E215	-	# Disable check for saved payment methods by sites (Category: Microsoft Edge (new version based on Chromium))
	E118	+	# Disable personalizing advertising, search, news and other services (Category: Microsoft Edge (new version based on Chromium))
	E218	-	# Disable personalizing advertising, search, news and other services (Category: Microsoft Edge (new version based on Chromium))
	E107	-	# Disable automatic completion of web addresses in address bar (Category: Microsoft Edge (new version based on Chromium))
	E207	-	# Disable automatic completion of web addresses in address bar (Category: Microsoft Edge (new version based on Chromium))
	E111	+	# Disable user feedback in toolbar (Category: Microsoft Edge (new version based on Chromium))
	E211	+	# Disable user feedback in toolbar (Category: Microsoft Edge (new version based on Chromium))
	E112	+	# Disable storing and autocompleting of credit card data on websites (Category: Microsoft Edge (new version based on Chromium))
	E212	+	# Disable storing and autocompleting of credit card data on websites (Category: Microsoft Edge (new version based on Chromium))
	E109	+	# Disable form suggestions (Category: Microsoft Edge (new version based on Chromium))
	E209	+	# Disable form suggestions (Category: Microsoft Edge (new version based on Chromium))
	E121	-	# Disable suggestions from local providers (Category: Microsoft Edge (new version based on Chromium))
	E221	-	# Disable suggestions from local providers (Category: Microsoft Edge (new version based on Chromium))
	E103	-	# Disable search and website suggestions (Category: Microsoft Edge (new version based on Chromium))
	E203	-	# Disable search and website suggestions (Category: Microsoft Edge (new version based on Chromium))
	E123	+	# Disable shopping assistant in Microsoft Edge (Category: Microsoft Edge (new version based on Chromium))
	E223	+	# Disable shopping assistant in Microsoft Edge (new version based on Chromium))
	E124	+	# Disable Edge bar (Category: Microsoft Edge (new version based on Chromium))
	E224	+	# Disable Edge bar (Category: Microsoft Edge (new version based on Chromium))
	E128	+	# Disable Sidebar in Microsoft Edge (Category: Microsoft Edge (new version based on Chromium))
	E228	+	# Disable Sidebar in Microsoft Edge (Category: Microsoft Edge (new version based on Chromium))
	E129	-	# Disable the Microsoft Account Sign-In Button (Category: Microsoft Edge (new version based on Chromium))
	E229	-	# Disable the Microsoft Account Sign-In Button (Category: Microsoft Edge (new version based on Chromium))
	E130	-	# Disable Enhanced Spell Checking (Category: Microsoft Edge (new version based on Chromium))
	E230	-	# Disable Enhanced Spell Checking (Category: Microsoft Edge (new version based on Chromium))
	E119	-	# Disable use of web service to resolve navigation errors (Category: Microsoft Edge (new version based on Chromium))
	E219	-	# Disable use of web service to resolve navigation errors (Category: Microsoft Edge (new version based on Chromium))
	E120	-	# Disable suggestion of similar sites when website cannot be found (Category: Microsoft Edge (new version based on Chromium))
	E220	-	# Disable suggestion of similar sites when website cannot be found (Category: Microsoft Edge (new version based on Chromium))
	E122	-	# Disable preload of pages for faster browsing and searching (Category: Microsoft Edge (new version based on Chromium))
	E222	-	# Disable preload of pages for faster browsing and searching (Category: Microsoft Edge (new version based on Chromium))
	E125	+	# Disable saving passwords for websites (Category: Microsoft Edge (new version based on Chromium))
	E225	-	# Disable saving passwords for websites (Category: Microsoft Edge (new version based on Chromium))
	E126	-	# Disable site safety services for more information about a visited website (Category: Microsoft Edge (new version based on Chromium))
	E226	-	# Disable site safety services for more information about a visited website (Category: Microsoft Edge (new version based on Chromium))
	E131	-	# Disable automatic redirection from Internet Explorer to Microsoft Edge (Category: Microsoft Edge (new version based on Chromium))
	E106	-	# Disable SmartScreen Filter (Category: Microsoft Edge (new version based on Chromium))
	E206	-	# Disable SmartScreen Filter (Category: Microsoft Edge (new version based on Chromium))
	E127	-	# Disable typosquatting checker for site addresses (Category: Microsoft Edge (new version based on Chromium))
	E227	-	# Disable typosquatting checker for site addresses (Category: Microsoft Edge (new version based on Chromium))
	E001	+	# Disable tracking in the web (Category: Microsoft Edge (legacy version))
	E002	+	# Disable page prediction (Category: Microsoft Edge (legacy version))
	E003	+	# Disable search and website suggestions (Category: Microsoft Edge (legacy version))
	E008	+	# Disable Cortana in Microsoft Edge (Category: Microsoft Edge (legacy version))
	E007	+	# Disable automatic completion of web addresses in address bar (Category: Microsoft Edge (legacy version))
	E010	+	# Disable showing search history (Category: Microsoft Edge (legacy version))
	E011	+	# Disable user feedback in toolbar (Category: Microsoft Edge (legacy version))
	E012	+	# Disable storing and autocompleting of credit card data on websites (Category: Microsoft Edge (legacy version))
	E009	+	# Disable form suggestions (Category: Microsoft Edge (legacy version))
	E004	+	# Disable sites saving protected media licenses on my device (Category: Microsoft Edge (legacy version))
	E005	+	# Do not optimize web search results on the task bar for screen reader (Category: Microsoft Edge (legacy version))
	E013	-	# Disable Microsoft Edge launch in the background (Category: Microsoft Edge (legacy version))
	E014	-	# Disable loading the start and new tab pages in the background (Category: Microsoft Edge (legacy version))
	E006	-	# Disable SmartScreen Filter (Category: Microsoft Edge (legacy version))
	Y001	-	# Disable synchronization of all settings (Category: Synchronization of Windows Settings)
	Y002	-	# Disable synchronization of design settings (Category: Synchronization of Windows Settings)
	Y003	-	# Disable synchronization of browser settings (Category: Synchronization of Windows Settings)
	Y004	-	# Disable synchronization of credentials (passwords) (Category: Synchronization of Windows Settings)
	Y005	-	# Disable synchronization of language settings (Category: Synchronization of Windows Settings)
	Y006	-	# Disable synchronization of accessibility settings (Category: Synchronization of Windows Settings)
	Y007	-	# Disable synchronization of advanced Windows settings (Category: Synchronization of Windows Settings)
	C012	+	# Disable and reset Cortana (Category: Cortana (Personal Assistant))
	C002	+	# Disable Input Personalization (Category: Cortana (Personal Assistant))
	C013	+	# Disable online speech recognition (Category: Cortana (Personal Assistant))
	C007	+	# Cortana and search are disallowed to use location (Category: Cortana (Personal Assistant))
	C008	+	# Disable web search from Windows Desktop Search (Category: Cortana (Personal Assistant))
	C009	+	# Disable display web results in Search (Category: Cortana (Personal Assistant))
	C010	+	# Disable download and updates of speech recognition and speech synthesis models (Category: Cortana (Personal Assistant))
	C011	+	# Disable cloud search (Category: Cortana (Personal Assistant))
	C014	+	# Disable Cortana above lock screen (Category: Cortana (Personal Assistant))
	C015	+	# Disable the search highlights in the taskbar (Category: Cortana (Personal Assistant))
	C101	+	# Disable the Windows Copilot (Category: Windows AI)
	C201	+	# Disable the Windows Copilot (Category: Windows AI)
	C204	+	# Disable the provision of recall functionality to all users (Category: Windows AI)
	C205	+	# Disable the Image Creator in Microsoft Paint (Category: Windows AI)
	C102	+	# Disable the Copilot button from the taskbar (Category: Windows AI)
	C103	+	# Disable Windows Copilot+ Recall (Category: Windows AI)
	C203	+	# Disable Windows Copilot+ Recall (Category: Windows AI)
	C206	+	# Disable Cocreator in Microsoft Paint (Category: Windows AI)
	C207	+	# Disable AI-powered image fill in Microsoft Paint (Category: Windows AI)
	L001	-	# Disable functionality to locate the system (Category: Location Services)
	L003	-	# Disable scripting functionality to locate the system (Category: Location Services)
	L004	-	# Disable sensors for locating the system and its orientation (Category: Location Services)
	L005	-	# Disable Windows Geolocation Service (Category: Location Services)
	U001	+	# Disable application telemetry (Category: User Behavior)
	U004	+	# Disable diagnostic data from customizing user experiences for whole machine (Category: User Behavior)
	U005	+	# Disable the use of diagnostic data for a tailor-made user experience for current user (Category: User Behavior)
	U006	+	# Disable diagnostic log collection (Category: User Behavior)
	U007	+	# Disable downloading of OneSettings configuration settings (Category: User Behavior)
	W001	-	# Disable Windows Update via peer-to-peer (Category: Windows Update)
	W011	-	# Disable updates to the speech recognition and speech synthesis modules. (Category: Windows Update)
	W004	-	# Activate deferring of upgrades (Category: Windows Update)
	W005	-	# Disable automatic downloading manufacturers' apps and icons for devices (Category: Windows Update)
	W010	-	# Disable automatic driver updates through Windows Update (Category: Windows Update)
	W009	-	# Disable automatic app updates through Windows Update (Category: Windows Update)
	P017	-	# Disable Windows dynamic configuration and update rollouts (Category: Windows Update)
	W006	-	# Disable automatic Windows Updates (Category: Windows Update)
	W008	-	# Disable Windows Updates for other products (e.g. Microsoft Office) (Category: Windows Update)
	M006	+	# Disable occassionally showing app suggestions in Start menu (Category: Windows Explorer)
	M011	-	# Do not show recently opened items in Jump Lists on "Start" or the taskbar (Category: Windows Explorer)
	M010	-	# Disable ads in Windows Explorer/OneDrive (Category: Windows Explorer)
	O003	-	# Disable OneDrive access to network before login (Category: Windows Explorer)
	O001	-	# Disable Microsoft OneDrive (Category: Windows Explorer)
	S012	-	# Disable Microsoft SpyNet membership (Category: Microsoft Defender and Microsoft SpyNet)
	S013	-	# Disable submitting data samples to Microsoft (Category: Microsoft Defender and Microsoft SpyNet)
	S014	-	# Disable reporting of malware infection information (Category: Microsoft Defender and Microsoft SpyNet)
	K001	+	# Disable Windows Spotlight (Category: Lock Screen)
	K002	+	# Disable fun facts, tips, tricks, and more on your lock screen (Category: Lock Screen)
	K005	-	# Disable notifications on lock screen (Category: Lock Screen)
	D001	+	# Disable access to mobile devices (Category: Mobile Devices)
	D002	+	# Disable Phone Link app (Category: Mobile Devices)
	D003	+	# Disable showing suggestions for using mobile devices with Windows (Category: Mobile Devices)
	D104	+	# Disable connecting the PC to mobile devices (Category: Mobile Devices)
	M025	+	# Disable search with AI in search box (Category: Search)
	M003	+	# Disable extension of Windows search with Bing (Category: Search)
	M015	+	# Disable People icon in the taskbar (Category: Taskbar)
	M016	+	# Disable search box in task bar (Category: Taskbar)
	M017	+	# Disable "Meet now" in the task bar on this device (Category: Taskbar)
	M018	+	# Disable "Meet now" in the task bar for current user (Category: Taskbar)
	M019	+	# Disable news and interests in the task bar on this device (Category: Taskbar)
	M021	-	# Disable widgets in Windows Explorer (Category: Taskbar)
	M022	+	# Disable feedback reminders on this device (Category: Miscellaneous)
	M001	+	# Disable feedback reminders for current user (Category: Miscellaneous)
	M004	+	# Disable automatic installation of recommended Windows Store Apps (Category: Miscellaneous)
	M005	+	# Disable tips, tricks, and suggestions while using Windows (Category: Miscellaneous)
	M024	+	# Disable Windows Media Player Diagnostics (Category: Miscellaneous)
	M012	-	# Disable Key Management Service Online Activation (Category: Miscellaneous)
	M013	-	# Disable automatic download and update of map data (Category: Miscellaneous)
	M014	-	# Disable unsolicited network traffic on the offline maps settings page (Category: Miscellaneous)
	M026	-	# Disable remote assistance connections to this computer (Category: Miscellaneous)
	M027	+	# Disable remote connections to this computer (Category: Miscellaneous)
	M028	+	# Disable the desktop icon for information on "Windows Spotlight" (Category: Miscellaneous)
	N001	-	# Disable Network Connectivity Status Indicator (Category: Miscellaneous)
'@

	Set-Content -Path $ooCfg -Value $ooConfigContent -Encoding UTF8
	try {
		$ProgressPreference = 'SilentlyContinue'
		Invoke-WebRequest -Uri 'https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe' -OutFile $ooExe
		Start-Process -FilePath $ooExe -ArgumentList "`"$ooCfg`" /quiet" -Wait -WindowStyle Hidden | Out-Null
	} catch { Write-ProgressLog "OOSU10 apply failed: $_" 'WARN' }
	finally {
		if (Test-Path $ooCfg) { Remove-Item $ooCfg -Force }
		if (Test-Path $ooExe) { Remove-Item $ooExe -Force }
		$ProgressPreference = 'Continue'
	}
}

function Set-LocationServices {
	Write-ProgressLog "Setting location services"
	Start-Process "ms-settings:privacy-location"
	$root = [System.Windows.Automation.AutomationElement]::RootElement
	$settings = Get-UIAElement -Root $root -Name "Settings"
	if ($settings) {
		$toggle = Get-UIAElement -Root $settings -AutomationId "DialogToggle" -Name "Location services"
		if ($toggle) {
			$togglePattern = $null
			if ($toggle.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
				if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
					Invoke-UIAElement -Element $toggle
					Start-Sleep -Milliseconds 200
				}
			}
		}

		$toggle = Get-UIAElement -Root $settings -AutomationId "SystemSettings_CapabilityAccess_Location_UserGlobal_ToggleSwitch" -Name "Let apps access your location"
		if ($toggle) {
			$togglePattern = $null
			if ($toggle.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
				if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
					Invoke-UIAElement -Element $toggle
					Start-Sleep -Milliseconds 200
				}
			}
		}

		$toggle = Get-UIAElement -Root $settings -AutomationId "SystemSettings_CapabilityAccess_Location_ClassicGlobal_ToggleSwitch" -Name "Let desktop apps access your location"
		if ($toggle) {
			$togglePattern = $null
			if ($toggle.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
				if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
					Invoke-UIAElement -Element $toggle
					Start-Sleep -Milliseconds 200
				}
			}
		}

		Close-SettingsWindow -Window $settings
	}
}

function Set-DateTime {
	Write-ProgressLog "Setting date and time"
	Start-Process "ms-settings:dateandtime"
	$root = [System.Windows.Automation.AutomationElement]::RootElement
	$settings = Get-UIAElement -Root $root -Name "Settings"
	if ($settings) {
		$toggles = @(
			"SystemSettings_DateTime_IsTimeSetAutomaticallyEnabled_ToggleSwitch",
			"SystemSettings_DateTime_IsTimeZoneSetAutomaticallyEnabled_ToggleSwitch"
		)
		foreach ($id in $toggles) {
			$toggle = Get-UIAElement -Root $settings -AutomationId $id
			if ($toggle) {
				$togglePattern = $null
				if ($toggle.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
					if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
						Invoke-UIAElement -Element $toggle
						Start-Sleep -Milliseconds 200
					}
				}
			}
		}
		Close-SettingsWindow -Window $settings
	}
}

function Set-RegistryTweaks {
	Write-ProgressLog "Applying registry and system tweaks"
	# Theme color prevalence off
	Set-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'ColorPrevalence' -Value 0
	Set-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Windows\DWM' -Name 'ColorPrevalence' -Value 0
	# Explorer start to OneDrive
	Set-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'LaunchTo' -Value 4
	# Hibernate and menu
	powercfg.exe /hibernate on
	$flyout = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings'
	Set-RegistryValueSafe -Path $flyout -Name 'ShowHibernateOption' -Value 1
	# Disable Bing search and web results
	$SearchPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
	Set-RegistryValueSafe -Path $SearchPath -Name 'BingSearchEnabled' -Value 0
	Set-RegistryValueSafe -Path $SearchPath -Name 'CortanaConsent' -Value 0
	$ExplorerPolicyPath = 'HKCU:\Software\Policies\Microsoft\Windows\Explorer'
	Set-RegistryValueSafe -Path $ExplorerPolicyPath -Name 'DisableSearchBoxSuggestions' -Value 1
	$WindowsSearchPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
	Set-RegistryValueSafe -Path $WindowsSearchPolicyPath -Name 'ConnectedSearchUseWeb' -Value 0
	Set-RegistryValueSafe -Path $WindowsSearchPolicyPath -Name 'AllowSearchToUseLocation' -Value 0
	$SearchSettingsPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings'
	Set-RegistryValueSafe -Path $SearchSettingsPath -Name 'IsDynamicSearchBoxEnabled' -Value 0
	# Ads/tips
	$CDMPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
	Set-RegistryValueSafe -Path $CDMPath -Name 'SubscribedContent-338387Enabled' -Value 0
	Set-RegistryValueSafe -Path $CDMPath -Name 'SubscribedContent-338389Enabled' -Value 0
	Set-RegistryValueSafe -Path $CDMPath -Name 'SubscribedContent-353694Enabled' -Value 0
	Set-RegistryValueSafe -Path $CDMPath -Name 'SubscribedContent-310093Enabled' -Value 0
	$ExplorerAdvanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'Start_TrackProgs' -Value 0
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'Start_TrackDocs' -Value 0
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'ShowSyncProviderNotifications' -Value 0
	# Clipboard history
	Set-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Clipboard' -Name 'EnableClipboardHistory' -Value 1
	# Execution policy
	try { Set-ExecutionPolicy Unrestricted -Force -ErrorAction Stop } catch {}
	# Alt+Tab tabs off
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'MultiTaskingAltTabFilter' -Value 3
	# Taskbar auto-hide
	try { &{ $p='HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3';$v=(Get-ItemProperty -Path $p).Settings;$v[8]=3;&Set-RegistryValueSafe -Path $p -Name Settings -Value $v -Type Binary;&Stop-Process -f -ProcessName explorer } } catch { Write-ProgressLog "Taskbar auto-hide failed: $_" 'WARN' }
	# End task on taskbar
	Set-RegistryValueSafe -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings' -Name 'TaskbarEndTask' -Value 1
	# Disable Recall via DISM
	Invoke-DismSafe -Arguments '/Online /Disable-Feature /FeatureName:Recall'
	# Remove Copilot using supported capability/package cmdlets (skip if absent)
	try {
		$copilotCap = Get-WindowsCapability -Online -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Microsoft.Windows.Copilot*' -and $_.State -eq 'Installed' }
		if ($copilotCap) {
			foreach ($cap in $copilotCap) { Remove-WindowsCapability -Online -Name $cap.Name -ErrorAction SilentlyContinue | Out-Null }
		} else {
			$copilotPkgs = Get-WindowsPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.PackageName -match 'Microsoft.Windows.Copilot' -and $_.ReleaseType -ne 'Feature Pack' }
			foreach ($pkg in $copilotPkgs) { Remove-WindowsPackage -Online -PackageName $pkg.PackageName -ErrorAction SilentlyContinue | Out-Null }
			if (-not $copilotCap -and ($null -eq $copilotPkgs -or $copilotPkgs.Count -eq 0)) { Write-ProgressLog 'Copilot not present; skipping removal' 'INFO' }
		}
	} catch { Write-ProgressLog "Copilot removal issues: $_" 'WARN' }
	# Kill Intel LMS service and remove
	try {
		Stop-Service -Name 'LMS' -Force -ErrorAction SilentlyContinue
		Set-Service -Name 'LMS' -StartupType Disabled -ErrorAction SilentlyContinue
		sc.exe delete LMS | Out-Null
		$lmsDriverPackages = Get-ChildItem -Path 'C:\Windows\System32\DriverStore\FileRepository' -Recurse -Filter 'lms.inf*' -ErrorAction SilentlyContinue
		foreach ($package in $lmsDriverPackages) { pnputil /delete-driver $($package.Name) /uninstall /force | Out-Null }
		$programFilesDirs = @('C:\Program Files','C:\Program Files (x86)')
		foreach ($dir in $programFilesDirs) {
			Get-ChildItem -Path $dir -Recurse -Filter 'LMS.exe' -ErrorAction SilentlyContinue | ForEach-Object {
				& icacls $_.FullName /grant Administrators:F /T /C /Q
				& takeown /F $_.FullName /A /R /D Y
				Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
			}
		}
	} catch { Write-ProgressLog "LMS removal issues: $_" 'WARN' }
	# Hibernation defaults
	Invoke-External -FilePath powercfg -Arguments '/hibernate on'
	Invoke-External -FilePath powercfg -Arguments '/change standby-timeout-ac 60'
	Invoke-External -FilePath powercfg -Arguments '/change standby-timeout-dc 60'
	Invoke-External -FilePath powercfg -Arguments '/change monitor-timeout-ac 10'
	Invoke-External -FilePath powercfg -Arguments '/change monitor-timeout-dc 1'
	# WSL/Sandbox features
	Invoke-DismSafe -Arguments '/online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart'
	Invoke-DismSafe -Arguments '/online /enable-feature /featurename:Containers-DisposableClientVM /all /norestart'
	Invoke-DismSafe -Arguments '/online /enable-feature /featurename:VirtualMachinePlatform /all /norestart'
	# Windows Update deferrals
	$UpdatePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
	Set-RegistryValueSafe -Path $UpdatePath -Name 'DeferFeatureUpdates' -Value 1
	Set-RegistryValueSafe -Path $UpdatePath -Name 'DeferFeatureUpdatesPeriodInDays' -Value 730
	Set-RegistryValueSafe -Path $UpdatePath -Name 'DeferQualityUpdates' -Value 0
	Set-RegistryValueSafe -Path $UpdatePath -Name 'DeferQualityUpdatesPeriodInDays' -Value 0
	# UI/taskbar preferences
	Set-RegistryValueSafe -Path $SearchPath -Name 'DisableSearchBoxSuggestions' -Value 1
	Set-RegistryValueSafe -Path $SearchPath -Name 'SearchboxTaskbarMode' -Value 0
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'ShowTaskViewButton' -Value 0
	try { Set-ItemProperty -Path $ExplorerAdvanced -Name 'TaskbarDa' -Value 0 -Force -ErrorAction Stop } catch { Write-ProgressLog 'TaskbarDa write blocked; using policy fallback' 'WARN' }
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'TaskbarEndTask' -Value 1
	# Policy fallbacks to disable Widgets/News if HKCU write is blocked
	Set-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh' -Name 'AllowNewsAndInterests' -Value 0
	Set-RegistryValueSafe -Path 'HKCU:\SOFTWARE\Policies\Microsoft\Dsh' -Name 'AllowNewsAndInterests' -Value 0
	$GalleryKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace_41040327\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}'
	if (Test-Path $GalleryKey) { Remove-Item -Path $GalleryKey -Force -Recurse -ErrorAction SilentlyContinue }
	# Privacy/Telemetry
	Set-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -Name 'AllowTelemetry' -Value 0
	Set-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent' -Name 'DisableWindowsConsumerFeatures' -Value 1
	Set-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config' -Name 'AutoConnectAllowedOEM' -Value 0
	# AI policies
	Set-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI' -Name 'DisableAIDataAnalysis' -Value 1
	Set-RegistryValueSafe -Path 'HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1
	Set-RegistryValueSafe -Path $ExplorerAdvanced -Name 'ShowCopilotButton' -Value 0
	# Power button to hibernate
	try {
		$scheme = [string](powercfg /getactivescheme).Split()[3]
		Invoke-External -FilePath powercfg -Arguments "/setacvalueindex $scheme SUB_BUTTONS PBUTTONACTION 2"
		Invoke-External -FilePath powercfg -Arguments "/setdcvalueindex $scheme SUB_BUTTONS PBUTTONACTION 2"
	} catch {}
	# WinUtil runner
	try {
		$tempDir = $env:TEMP
		$jsonPath = "$tempDir\winutil_config.json"
		$runnerPath = "$tempDir\winutil_direct_runner.ps1"
		$jsonContent = @"
	{
		"WPFTweaks":  [
						"WPFTweaksWifi",
						"WPFTweaksServices",
						"WPFToggleBingSearch",
						"WPFToggleTaskbarSearch",
						"WPFToggleTaskView",
						"WPFToggleTaskbarWidgets",
						"WPFTweaksConsumerFeatures",
						"WPFTweaksTele",
						"WPFTweaksRemoveGallery",
						"WPFTweaksEndTaskOnTaskbar",
						"WPFTweaksRemoveCopilot",
						"WPFTweaksLaptopHibernation",
						"WPFTweaksRecallOff"
					],
		"Install":  [],
		"WPFInstall": [],
		"WPFFeature": [
						"WPFFeaturesSandbox",
						"WPFFeaturewsl"
					]
	}
	"@
		Set-Content -Path $jsonPath -Value $jsonContent -Encoding UTF8
		$runnerContent = @"
		`$OutputEncoding = [System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8
		Write-Host "Starting Download..."
		`$scriptContent = Invoke-RestMethod "https://christitus.com/win"
		`$scriptBlock = [Scriptblock]::Create(`$scriptContent)
		Write-Host "Running Tweaks..."
		& `$scriptBlock -Config "$jsonPath" -Run
"@
		Set-Content -Path $runnerPath -Value $runnerContent -Encoding UTF8
		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = 'powershell.exe'
		$pinfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$runnerPath`""
		$pinfo.RedirectStandardOutput = $true
		$pinfo.UseShellExecute = $false
		$pinfo.CreateNoWindow = $true
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$p.Start() | Out-Null
		while (-not $p.StandardOutput.EndOfStream) {
			$line = $p.StandardOutput.ReadLine()
			if ($line) { Write-ProgressLog "WinUtil: $line" }
			if ($line -match 'Tweaks\s+are\s+Finished') {
				Start-Sleep -Seconds 3
				try { $p.Kill() } catch {}
				break
			}
		}
	} catch { Write-ProgressLog "WinUtil failed: $_" 'WARN' }
	finally {
		if ($p -and -not $p.HasExited) { try { $p.Kill() } catch {} }
		if (Test-Path $jsonPath) { Remove-Item $jsonPath -Force }
		if (Test-Path $runnerPath) { Remove-Item $runnerPath -Force }
	}
}

function Set-ThemeAccess {
	$acl = Get-Acl "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
	$rule = New-Object System.Security.AccessControl.RegistryAccessRule ("ALL APPLICATION PACKAGES","ReadKey","Allow")
	$acl.SetAccessRule($rule)
	Set-Acl "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" $acl
}

function Set-PowerSleepPolicy {

	function Get-PowerGUIDMap {
		Write-ProgressLog "Mapping system power GUIDs..."
		$map = @{ Subgroups = @{}; Settings = @{} }
		
		# /QH ensures we see everything, even if the manufacturer tried to hide it
		$output = powercfg /qh
		
		foreach ($line in $output) {
			if ($line -match 'Subgroup GUID:\s+([a-f0-9-]+)\s+\((.+)\)') {
				$map.Subgroups[$matches[2].Trim()] = $matches[1]
			}
			elseif ($line -match 'Power Setting GUID:\s+([a-f0-9-]+)\s+\((.+)\)') {
				$map.Settings[$matches[2].Trim()] = $matches[1]
			}
		}
		return $map
	}

    # 1. Ensure Hibernate is enabled and ready
    powercfg /hibernate on
    powercfg /h /type full

    # 2. Map current system GUIDs
    $PowerMap = Get-PowerGUIDMap
    $activeScheme = (powercfg /getactivescheme).Split()[3]
    Write-ProgressLog "Targeting Active Scheme: $activeScheme"

    # 3. Define the friendly names we are looking for
    # (Matches standard Windows 10/11 naming conventions)
    $sub_video_name   = "Display"
    $sub_sleep_name   = "Sleep"
    $sub_buttons_name = "Power buttons and lid"

    $set_display_name = "Turn off display after"
    $set_sleep_name   = "Sleep after"
    $set_hiber_name   = "Hibernate after"
    $set_pwrbtn_name  = "Power button action"
    $set_slpbtn_name  = "Sleep button action"
    $set_lid_name     = "Lid close action"

    # Helper to set values only if the GUID was successfully found
    $SetVal = {
        param($Mode, $SubName, $SetName, $Value)
        $subGuid = $PowerMap.Subgroups[$SubName]
        $setGuid = $PowerMap.Settings[$SetName]

        if ($subGuid -and $setGuid) {
            if ($Mode -eq "AC") {
                powercfg /setacvalueindex $activeScheme $subGuid $setGuid $Value
            } else {
                powercfg /setdcvalueindex $activeScheme $subGuid $setGuid $Value
            }
        } else {
            Write-ProgressLog "Warning: Could not find GUID for '$SetName' under '$SubName'" -Level "WARN"
        }
    }

    Write-ProgressLog "Applying Timeouts..."
    # Display: 2m (120s)
    &$SetVal "AC" $sub_video_name $set_display_name 120
    &$SetVal "DC" $sub_video_name $set_display_name 120

    # Sleep: Never (0s)
    &$SetVal "AC" $sub_sleep_name $set_sleep_name 0
    &$SetVal "DC" $sub_sleep_name $set_sleep_name 0

    # Hibernate: AC 10m (600s), DC 5m (300s)
    &$SetVal "AC" $sub_sleep_name $set_hiber_name 600
    &$SetVal "DC" $sub_sleep_name $set_hiber_name 300

    Write-ProgressLog "Applying Button Actions..."
    # Actions: 0=Nothing, 1=Sleep, 2=Hibernate
    &$SetVal "AC" $sub_buttons_name $set_pwrbtn_name 2
    &$SetVal "DC" $sub_buttons_name $set_pwrbtn_name 2

    &$SetVal "AC" $sub_buttons_name $set_slpbtn_name 2
    &$SetVal "DC" $sub_buttons_name $set_slpbtn_name 2

    &$SetVal "AC" $sub_buttons_name $set_lid_name 0
    &$SetVal "DC" $sub_buttons_name $set_lid_name 0

    # 4. Activate changes
    powercfg /setactive $activeScheme
    Write-ProgressLog "All policies applied successfully."
}

function Install-WingetPackages {
	Write-ProgressLog "Installing winget packages"
	Test-WingetAvailable
	$packages = @('Microsoft.Office','Microsoft.VisualStudioCode','voidtools.Everything','9NKSQGP7F2NH','Canva.Affinity','Microsoft.Git','Python.Python.3.12','LocalSend.LocalSend', 'namazso.SecureUXTheme', 'Wondershare.PDFelement.10','Microsoft.PowerToys','AdGuard.AdGuard')
	if ($PSVersionTable.PSVersion.Major -ge 7) {
		$packages | ForEach-Object -Parallel {
			winget install -e --id $_ --accept-package-agreements --accept-source-agreements
		} -ThrottleLimit 4
	} else {
		$jobs = New-Object System.Collections.ArrayList
		foreach ($pkg in $packages) {
			$jobs.Add((Start-Job -ScriptBlock { param($p) winget install -e --id $p --accept-package-agreements --accept-source-agreements } -ArgumentList $pkg)) | Out-Null
		}
		Wait-JobWithOutput -Jobs $jobs
	}

	Invoke-MinimizeAllWindows
}

function Set-VSCodeContext {
	Write-ProgressLog "Adding VS Code context menu entries"
	$CodePath = $null
	$possiblePaths = @(
		"$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
		"$env:ProgramFiles\Microsoft VS Code\Code.exe",
		"${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe"
	)
	foreach ($path in $possiblePaths) { if (Test-Path $path) { $CodePath = $path; break } }
	if (-not $CodePath) {
		$codeCommand = Get-Command code -ErrorAction SilentlyContinue
		if ($codeCommand -and $codeCommand.Source) { $CodePath = $codeCommand.Source }
	}
	if (-not $CodePath) {
		$regPaths = @(
			'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
			'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
			'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
		)
		foreach ($regPath in $regPaths) {
			$vscodeEntry = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*Visual Studio Code*' } | Select-Object -First 1
			if ($vscodeEntry -and $vscodeEntry.InstallLocation) {
				$potentialPath = Join-Path $vscodeEntry.InstallLocation 'Code.exe'
				if (Test-Path $potentialPath) { $CodePath = $potentialPath; break }
			}
		}
	}
	if (-not $CodePath) { Write-ProgressLog 'VS Code not found; skipping context menu' 'WARN'; return }
	function Add-ContextMenuEntry {
		param([string]$Target,[string]$MenuText,[string]$IconPath,[string]$Command)
		try {
			$rootKey = [Microsoft.Win32.Registry]::CurrentUser
			$shellKey = $rootKey.CreateSubKey("Software\Classes\$Target\shell\VSCode")
			$shellKey.SetValue('', $MenuText, [Microsoft.Win32.RegistryValueKind]::String)
			$shellKey.SetValue('Icon', $IconPath, [Microsoft.Win32.RegistryValueKind]::String)
			$shellKey.Close()
			$cmdKey = $rootKey.CreateSubKey("Software\Classes\$Target\shell\VSCode\command")
			$cmdKey.SetValue('', $Command, [Microsoft.Win32.RegistryValueKind]::String)
			$cmdKey.Close()
		} catch { Write-ProgressLog "Context menu failed for ${Target}: $_" 'WARN' }
	}
	Add-ContextMenuEntry -Target '*' -MenuText 'Open with Code' -IconPath $CodePath -Command "`"$CodePath`" `"%1`""
	Add-ContextMenuEntry -Target 'Directory' -MenuText 'Open with Code' -IconPath $CodePath -Command "`"$CodePath`" `"%V`""
	Add-ContextMenuEntry -Target 'Directory\Background' -MenuText 'Open with Code' -IconPath $CodePath -Command "`"$CodePath`" `"%V`""
}

function Set-PowerToysConfig {
	Write-ProgressLog "Configuring PowerToys"
	$ptSettingsDir  = "$env:LOCALAPPDATA\Microsoft\PowerToys"
	$globalSettings = "$ptSettingsDir\settings.json"
	$globalBackup   = "$ptSettingsDir\backup.json"
	$qaSettingsDir  = "$ptSettingsDir\QuickAccent"
	$qaSettingsPath = "$qaSettingsDir\settings.json"
	$kmSettingsDir  = "$ptSettingsDir\Keyboard Manager"
	$kmSettingsPath = "$kmSettingsDir\default.json"
	$exePath        = "$env:LOCALAPPDATA\PowerToys\PowerToys.exe"
	taskkill /F /IM PowerToys* /T 2>$null
	Start-Sleep -Seconds 3
	if (Test-Path $globalSettings) {
		$json = Get-Content -Path $globalSettings -Raw | ConvertFrom-Json
		$json.enabled.QuickAccent = $true
		$json.enabled."Keyboard Manager" = $true
		$json.enabled.CmdPal = $false
		Save-PTJson $globalSettings $json
		if (Test-Path $globalBackup) { Save-PTJson $globalBackup $json }
	}
	if (!(Test-Path $qaSettingsDir)) { New-Item -ItemType Directory -Path $qaSettingsDir | Out-Null }
	$qaConfig = @{
		properties = @{
			activation_key = 2
			do_not_activate_on_game_mode = $true
			toolbar_position = @{ value = 'Top center' }
			input_time_ms = @{ value = 300 }
			selected_lang = @{ value = 'SP' }
			excluded_apps = @{ value = '' }
			show_description = $false
			sort_by_usage_frequency = $true
			start_selection_from_the_left = $true
		}
		name = 'QuickAccent'
		version = '1.0'
	}
	Save-PTJson $qaSettingsPath $qaConfig
	if (!(Test-Path $kmSettingsDir)) { New-Item -ItemType Directory -Path $kmSettingsDir | Out-Null }
	$kmConfig = @{
		remapKeys = @{ inProcess = @() }
		remapKeysToText = @{ inProcess = @() }
		remapShortcuts = @{
			global = @(
				@{
					originalKeys = "162;160;68"
					exactMatch = $false
					runProgramElevationLevel = 0
					operationType = 1
					runProgramAlreadyRunningAction = 0
					runProgramStartWindowType = 1
					runProgramFilePath = "C:\Windows\System32\cmd.exe"
					runProgramArgs = "/c powershell -ExecutionPolicy Bypass -File C:\Windows\Resources\Themes\ToggleTheme.ps1"
					runProgramStartInDir = ""
					unicodeText = "*Unsupported*"
				}
			)
			appSpecific = @()
		}
		remapShortcutsToText = @{ global = @(); appSpecific = @() }
	}
	Save-PTJson $kmSettingsPath $kmConfig
	if (Test-Path $exePath) { Start-Process -FilePath $exePath } else { Write-ProgressLog 'PowerToys.exe not found' 'WARN' }
}

function Disable-StartMenuSuggestions {
	Write-ProgressLog "Disabling Start Menu suggestions"
	
	try {
		# Disable Iris Recommendations in Start Menu
		$path1 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
		if (-not (Test-Path $path1)) {
			New-Item -Path $path1 -Force | Out-Null
		}
		Set-RegistryValueSafe -Path $path1 -Name "Start_IrisRecommendations" -Value 0 -Type DWord -Force
		
		# Hide Recommended Personalized Sites (HKCU)
		$path2 = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
		if (-not (Test-Path $path2)) {
			New-Item -Path $path2 -Force | Out-Null
		}
		Set-RegistryValueSafe -Path $path2 -Name "HideRecommendedPersonalizedSites" -Value 1 -Type DWord -Force
		Set-RegistryValueSafe -Path $path2 -Name "DisableSearchBoxSuggestions" -Value 1 -Type DWord -Force
		
		# Hide Recommended Personalized Sites (HKLM - requires admin)
		$path3 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
		if (-not (Test-Path $path3)) {
			New-Item -Path $path3 -Force | Out-Null
		}
		Set-RegistryValueSafe -Path $path3 -Name "HideRecommendedPersonalizedSites" -Value 1 -Type DWord -Force
		
	} catch {
		Write-ProgressLog "Failed to disable Start Menu suggestions: $_" 'WARN'
	}
}

function Get-AndLockFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    process {
        try {
            $oldPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            # --- 1. FORCE REMOVAL OF EXISTING LOCKED FILE ---
            if (Test-Path $DestinationPath) {
                Write-ProgressLog "Existing locked file detected. Forcing deletion..."
                
                # Take ownership back (necessary if 'Everyone:Deny' is in place)
                # /F = File, /A = Give ownership to Administrators group
                $null = takeown.exe /F $DestinationPath /A

                # Grant Administrators full control and reset inheritance to wipe Deny rules
                $null = icacls.exe $DestinationPath /reset
                $null = icacls.exe $DestinationPath /grant "*S-1-5-32-544:F" # S-1-5-32-544 is Administrators

                # Clear the Read-Only attribute
                $existingFile = Get-Item $DestinationPath -Force
                $existingFile.Attributes = 'Normal'

                # Delete the file
                Remove-Item -Path $DestinationPath -Force -ErrorAction Stop
                Write-ProgressLog "Previous file removed successfully."
            }

            # --- 2. DOWNLOAD NEW FILE ---
            Write-ProgressLog "Downloading file..."
            Start-BitsTransfer -Source $Url -Destination $DestinationPath -ErrorAction Stop
            $ProgressPreference = $oldPreference

            # --- 3. APPLY LOCKDOWN ---
            $file = Get-Item -Path $DestinationPath
            $fullPath = $file.FullName

            # Set Read-Only attribute (Do this BEFORE applying the Deny ACL)
            Write-ProgressLog "Setting Read-Only attribute..."
            $file.IsReadOnly = $true

            Write-ProgressLog "Applying security lockdown (Deny Write)..."
            $acl = Get-Acl -Path $fullPath
            
            # Disable inheritance and remove all existing rules ($true, $false)
            $acl.SetAccessRuleProtection($true, $false)

            # Define 'Everyone' SID (S-1-1-0)
            $everyone = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::WorldSid, $null)

            # Rule: Allow Read (Required so the system can still use the file)
            $allowRead = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $everyone,
                [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            $acl.AddAccessRule($allowRead)

            # Rule: Deny Write, Modify, Delete, and Attribute changes
            # This prevents any principal (even System/Admin) from changing the file
            $denyRights = [System.Security.AccessControl.FileSystemRights]"Write, Modify, Delete, AppendData, WriteAttributes, WriteExtendedAttributes, DeleteSubdirectoriesAndFiles"
            $denyWrite = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $everyone,
                $denyRights,
                [System.Security.AccessControl.AccessControlType]::Deny
            )
            $acl.AddAccessRule($denyWrite)

            # Apply the finished ACL
            Set-Acl -Path $fullPath -AclObject $acl

            Write-ProgressLog "Process Complete. File is now immutable." -Level "SUCCESS"
        }
        catch {
            if ($null -ne $oldPreference) { $ProgressPreference = $oldPreference }
            Write-Error "Failed to process file. Error: $($_.Exception.Message)"
        }
    }
}

function Disable-StoreSearchResults {
    try {
        Write-ProgressLog "Disabling Microsoft Store search results..."
        $storeUrl = "https://github.com/mcAlex42/WinSetup/raw/refs/heads/main/Installers/store.db"
        $storePkg = Get-AppxPackage -Name Microsoft.WindowsStore | Select-Object -First 1
        $pfn = if ($storePkg) { $storePkg.PackageFamilyName } else { "Microsoft.WindowsStore_8wekyb3d8bbwe" }
        $destinationPath = "$env:LOCALAPPDATA\Packages\$pfn\LocalState\store.db"

        Get-AndLockFile -Url $storeUrl -DestinationPath $destinationPath
    }
    catch {
        Write-ProgressLog "Failed to disable Microsoft Store search results. Error: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Set-StartupApps {
	Write-ProgressLog "Disabling startup applications"
	function Set-RegistryValueLocal {
		param([string]$Path,[string]$Name,[string]$Type,[Object]$Value)
		if ($Path.StartsWith('HKCU')) { $Path = $Path.Replace('HKCU','HKCU:') }
		if ($Path.StartsWith('HKLM')) { $Path = $Path.Replace('HKLM','HKLM:') }
		if (Test-Path $Path) {
			try { New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force -ErrorAction Stop | Out-Null } catch {}
		}
	}
	$CurrentTime = [DateTime]::Now.ToFileTime()
	$waPath = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\2defd21c-0b9e-4e4e-873a-2a68c47d7da5'
	Set-RegistryValueLocal -Path $waPath -Name 'State' -Type DWord -Value 1; Set-RegistryValueLocal -Path $waPath -Name 'LastDisabledTime' -Type QWord -Value $CurrentTime
	$cdPath = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\MicrosoftWindows.CrossDevice_cw5n1h2txyewy\CrossDevice.Start'
	Set-RegistryValueLocal -Path $cdPath -Name 'State' -Type DWord -Value 1; Set-RegistryValueLocal -Path $cdPath -Name 'LastDisabledTime' -Type QWord -Value $CurrentTime
	$teamsPath = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\MSTeams_8wekyb3d8bbwe\TeamsTfwStartupTask'
	Set-RegistryValueLocal -Path $teamsPath -Name 'State' -Type DWord -Value 1; Set-RegistryValueLocal -Path $teamsPath -Name 'LastDisabledTime' -Type QWord -Value $CurrentTime
	$intelPath = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\AppUp.IntelConnectivityPerformanceSuite_8j3eq9eme6ctt\ICMTask'
	Set-RegistryValueLocal -Path $intelPath -Name 'State' -Type DWord -Value 1; Set-RegistryValueLocal -Path $intelPath -Name 'LastDisabledTime' -Type QWord -Value $CurrentTime
	$cmdPalPath = 'HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\SystemAppData\Microsoft.CommandPalette_8wekyb3d8bbwe\CmdPalStartup'
	Set-RegistryValueLocal -Path $cmdPalPath -Name 'State' -Type DWord -Value 1; Set-RegistryValueLocal -Path $cmdPalPath -Name 'LastDisabledTime' -Type QWord -Value $CurrentTime
	$binData1 = [byte[]](0x03,0x00,0x00,0x00,0x0B,0x6F,0xDB,0x94,0x42,0x74,0xDC,0x01)
	Set-RegistryValueLocal -Path 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' -Name 'Wondershare PEScreenshot.lnk' -Type Binary -Value $binData1
	$binData2 = [byte[]](0x03,0x00,0x00,0x00,0x4A,0x2D,0xDD,0x95,0x42,0x74,0xDC,0x01)
	Set-RegistryValueLocal -Path 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' -Name 'Wondershare PEToolbox.lnk' -Type Binary -Value $binData2
	$binData3 = [byte[]](0x03,0x00,0x00,0x00,0x95,0x61,0xEB,0x96,0x42,0x74,0xDC,0x01)
	Set-RegistryValueLocal -Path 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder' -Name 'Send to OneNote.lnk' -Type Binary -Value $binData3
	$binData4 = [byte[]](0x03,0x00,0x00,0x00,0x9D,0xEF,0xB6,0x9C,0x42,0x74,0xDC,0x01)
	Set-RegistryValueLocal -Path 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' -Name 'Logitech Download Assistant' -Type Binary -Value $binData4
	$binDataEverything = [byte[]](0x03,0x00,0x00,0x00,0x3C,0xBB,0xC0,0x78,0x60,0x74,0xDC,0x01)
	Set-RegistryValueLocal -Path 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' -Name 'Everything' -Type Binary -Value $binDataEverything
}

function Set-NightLight {
	Write-ProgressLog "Setting night light"
	Start-Process 'ms-settings:nightlight'
	$root = [System.Windows.Automation.AutomationElement]::RootElement
	$settings = Get-UIAElement -Root $root -Name "Settings"
	if ($settings) {
		$slider = Get-UIAElement -Root $settings -AutomationId 'SystemSettings_Display_BlueLight_ColorTemperature_Slider'
		if ($slider) {
			try {
				$rangeValuePattern = $null
				if ($slider.TryGetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern, [ref]$rangeValuePattern)) {
					$rangeValuePattern.SetValue(80)
				}
			} catch { Write-ProgressLog "Night light slider failed: $_" 'WARN' }
		}
		$toggleSwitch = Get-UIAElement -Root $settings -AutomationId 'SystemSettings_Display_BlueLight_AutomaticOnSchedule_ToggleSwitch'
		if ($toggleSwitch) {
			try {
				$togglePattern = $null
				if ($toggleSwitch.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
					if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
						$togglePattern.Toggle()
						Start-Sleep -Milliseconds 200
					}
				}
			} catch { Write-ProgressLog "Night light toggle failed: $_" 'WARN' }
		}
		Close-SettingsWindow -Window $settings
	}
}

function Set-DynamicRefresh {
    Write-ProgressLog "Setting dynamic refresh"
    Start-Process 'ms-settings:display'
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $settings = Get-UIAElement -Root $root -Name "Settings"
    
    if ($settings) {
        # 1. Navigate to Advanced Display
        $advContainer = Get-UIAElement -Root $settings -AutomationId 'SystemSettings_Display_AdvancedDisplaySettings_ButtonEntityItem'
        if ($advContainer) {
            try { ($advContainer.GetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern)).ScrollIntoView() } catch {}
            $buttonCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Button)
            $childButton = $advContainer.FindFirst([System.Windows.Automation.TreeScope]::Children,$buttonCondition)
            if ($childButton) { Invoke-UIAElement -Element $childButton } else { try { $advContainer.SetFocus(); [System.Windows.Forms.SendKeys]::SendWait('{TAB}'); [System.Windows.Forms.SendKeys]::SendWait('{ENTER}') } catch {} }
        } else { 
            Write-ProgressLog "Advanced Display container not found" 'WARN'
            Close-SettingsWindow -Window $settings
            return 
        }
        
        # 2. Find and Expand the Refresh Rate ComboBox
        $comboBox = Get-UIAElement -Root $settings -AutomationId 'SystemSettings_Display_AdvancedDisplaySettingsRefreshRate_ComboBox'
        if ($comboBox) {
            try {
                $expandPattern = $null
                if ($comboBox.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandPattern)) {
                    $expandPattern.Expand()
                    Start-Sleep -Milliseconds 800 # Give UI time to populate items
                    
                    # Fetch list items
                    $listItems = $comboBox.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition) | 
                                 Where-Object { $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::ListItem }
                    
                    if ($listItems.Count -le 1) {
                        Write-ProgressLog "Only $($listItems.Count) refresh rate option(s) found. No selection necessary."
                        Close-SettingsWindow -Window $settings
                        return
                    }
                    Write-ProgressLog "Found $($listItems.Count) refresh rates. Selecting highest..."

                    # 3. Find and select the highest refresh rate
                    $highestItem = $null
                    $highestValue = -1
                    foreach ($item in $listItems) {
                        if ($item.Current.Name -match '(\d+(\.\d+)?)') {
                            $val = [double]$matches[1]
                            if ($val -gt $highestValue) {
                                $highestValue = $val
                                $highestItem = $item
                            }
                        }
                    }
                    if ($highestItem) {
                        Write-ProgressLog "Selecting refresh rate: $($highestItem.Current.Name)"
                        $selectPattern = $null
                        if ($highestItem.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectPattern)) {
                            $selectPattern.Select()
                            Start-Sleep -Milliseconds 800
                        }
                    }
                }
            } catch { 
                Write-ProgressLog "Refresh combo interaction failed: $_" 'WARN' 
            }
        }

        # 4. Handle "Keep changes" prompt if it appears
        $primaryButton = Get-UIAElement -Root $settings -AutomationId 'PrimaryButton' -TimeoutSeconds 3
        if ($primaryButton) { 
            Write-ProgressLog "Confirming refresh rate change..."
            Invoke-UIAElement -Element $primaryButton
            Start-Sleep -Milliseconds 200
        }

        # 5. Handle Dynamic Refresh Toggle
        $toggleSwitch = Get-UIAElement -Root $settings -AutomationId 'SystemSettings_Display_AdvancedDisplaySettingsDynamicRefreshRate_ToggleSwitch'
        if ($toggleSwitch -and $toggleSwitch.Current.IsEnabled) {
            try {
                $togglePattern = $null
                if ($toggleSwitch.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
                    if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
                        Write-ProgressLog "Enabling Dynamic Refresh Rate toggle..."
                        $togglePattern.Toggle()
                        Start-Sleep -Milliseconds 200
                    }
                }
            } catch { 
                Write-ProgressLog "Dynamic refresh toggle failed: $_" 'WARN' 
            }
        }
        
        Close-SettingsWindow -Window $settings
    }
}

function Install-RemoteInstaller {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url
    )

    # 1. Prepare file path
    $fileName = Split-Path -Leaf $Url
    $tempPath = Join-Path -Path $env:TEMP -ChildPath $fileName

    # Store original preference to restore it later
    $originalPreference = $ProgressPreference

    try {
        # 2. Disable progress bar for speed
        $ProgressPreference = 'SilentlyContinue'
        Start-BitsTransfer -Source $Url -Destination $tempPath -Description "Downloading $fileName" -Priority High
        
        # Restore preference immediately after download
        $ProgressPreference = $originalPreference
        Write-ProgressLog "Download completed: $tempPath"

        # 3. Execute the installer
        if (Test-Path $tempPath) {
            Write-ProgressLog "Executing installer..."
            $process = Start-Process -FilePath $tempPath -Wait -PassThru
            Write-ProgressLog "Installation process finished with exit code: $($process.ExitCode)"
        }
        else {
            throw "File not found at $tempPath after download."
        }
    }
    catch {
        Write-ProgressLog "An error occurred: $_" "ERROR"
    }
    finally {
        # Restore preference in case of early crash
        $ProgressPreference = $originalPreference
        
        # 4. Cleanup
        if (Test-Path $tempPath) {
            Remove-Item -Path $tempPath -Force
        }
    }
}

# --- Installer URLs ---
$cursorInstallerUrl = "https://github.com/mcAlex42/WinSetup/raw/refs/heads/main/Installers/CursorInstall.exe"
$themeInstallerUrl = "https://github.com/mcAlex42/WinSetup/raw/refs/heads/main/Installers/ThemeInstaller.exe"
$darktitleInstallerUrl = "https://github.com/mcAlex42/WinSetup/raw/refs/heads/main/Installers/DarkTitleInstall.exe"

# Execution order ------------------------------------------------------
try {
	Invoke-MinimizeAllWindows
	Set-WindowsTheme
	Write-ProgressLog "Setting custom cursor..."
	Install-RemoteInstaller -Url $cursorInstallerUrl
	Write-ProgressLog "Setting custom theme..."
	Install-RemoteInstaller -Url $themeInstallerUrl
	Write-ProgressLog "Setting dark title bars..."
	Install-RemoteInstaller -Url $darktitleInstallerUrl
	Start-Process "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\DarkTitle.exe" -WindowStyle Hidden
	Set-PrivacySettings
	Set-RegistryTweaks
	Set-ThemeAccess
	Install-WingetPackages
	Set-PowerSleepPolicy
	Set-VSCodeContext
	Set-PowerToysConfig
	Set-StartupApps
	Disable-StartMenuSuggestions
	Disable-StoreSearchResults
	Close-SettingsWindow -Window (Get-UIAElement -Root ([System.Windows.Automation.AutomationElement]::RootElement) -Name 'Settings' -TimeoutSeconds 2)
	Set-LocationServices
	Set-DateTime
	Set-DynamicRefresh
	Stop-Process -Name 'SystemSettings' -Force -ErrorAction SilentlyContinue
	Set-NightLight
	Stop-Process -Name 'SystemSettings' -Force -ErrorAction SilentlyContinue
	Write-ProgressLog "Setup complete" 'DONE'
} catch {
	Write-ProgressLog "Setup halted: $_" 'ERROR'
}

# Wait for user confirmation before closing
Write-Host "Press enter to exit..."
Read-Host