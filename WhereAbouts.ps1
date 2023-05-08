Param(
    [Switch]$Console = $false           #--[ Set to true to enable local console result display. Defaults to false ]--
)
<#------------------------------------------------------------------------------ 
         File Name : Whereabouts.ps1 
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com) 
                   : 
       Description : Programatically creates an email to send to the predetermined user or group for notification about where you are.
                   : Grabs Outlook email signature from current users profile.  Determines sender by current logged on user. 
                   : 
             Notes : Normal operation is with no command line options.  
                   : See end of script for detail about how to launch via shortcut. 
                   : 
         Arguments : Command line options for testing: 
                   : - "-console $true" will enable local console echo for troubleshooting
                   : - "-debug $true" will only email to the first recipient on the list
                   : 
          Warnings : None 
                   : 
             Legal : Public Domain. Modify and redistribute freely. No rights reserved. 
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED. 
                   : That being said, please let me know if you find bugs or improve the script. 
                   : 
           Credits : Code snippets and/or ideas came from many sources including but 
                   : not limited to the following: n/a 
                   : 
    Last Update by : Kenneth C. Mazie 
   Version History : v1.00 - 09-09-18 - Original 
    Change History : v2.00 - 10-18-20 - Added travel log tracking and PC console locking options.
                   : v3.00 - 10-27-20 - Added test mode option. 
                   : v3.10 - 11-13-20 - Added function to assume job of updating run data
                   : v4.00 - 11-13-20 - Added option to dynamically size form depending on number of sites in site list
                   : v5.00 - 11-25-20 - Removed descrete checkbox sections and replaced with dynamic creator based 
                   :                    on site list.
                   : v5.10 - 04-28-21 - Minor form layout adjustment.
                   : v5.20 - 03-30-22 - Added 0 mile logging, reordered variable location, adjusted some console message
                   :                    text messages for clarity                
                   : #>
                   $Script:ScriptVer = "5.20"    <#--[ Current version # used in script ]--
                   : 
------------------------------------------------------------------------------#>
<#PSScriptInfo 
.VERSION 5.20 
.GUID 
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr.com) 
.DESCRIPTION 
Programatically creates an email to send to the predetermined user or group for notification about where you are.
Grabs Outlook email signature from current users profile.  Determines sender by current logged on user.  Optionally
writes a travel log to users Documents folder and optionally will lock the PC.  Dynamically grows or shrinks the
form depending on the number of sites in the site table.
#>
#Requires -Version 5.1

Clear-Host 

#--[ For Testing ]-------------
#$Script:Console = $true
#$Script:Debug = $true
#------------------------------

#--[ Suppress Console ]-------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition ' 
[DllImport("Kernel32.dll")] 
public static extern IntPtr GetConsoleWindow(); 
 
[DllImport("user32.dll")] 
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow); 
' 
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
#------------------------------------------------------------------------------#
 
#--[ Runtime Variables ]----------------------------------------------------
$ErrorActionPreference = "silentlycontinue"
$Script:Icon = [System.Drawing.SystemIcons]::Information
$Script:ReportBody = ""
$Script:ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0] 
$Script:LaunchLocation = "$Env:AppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
$Script:ConfigFile = $PSScriptRoot+'\'+$Script:ScriptName+'.xml'
$Script:TravelLogFile = "TravelLog.csv"
$Script:TravelLogPath = "$Env:UserProfile\Documents\"
$Script:Validated = $False
#$DomainName = $env:USERDOMAIN      #--[ Pulls local domain as an alternate if the user leaves it out ]-------
$Script:Message = ""
[Int]$Flag = 0
$Script:Refresh = $False
$Script:RunData = ""
$Script:LogZero = $False #True             #--[ Switch to TRUE to register zero milage destinations in the log ]--
$UN = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name        #--[ Current User   ]--
$DN = $UN.split("\")[0]                                                     #--[ Current Domain ]--    
$Script:SenderEmail = $UN.split("\")[1]+"@"+$DN+".org"                      #--[ Correct this for email domain, .ORG, .COM, etc ]--
$Script:TitleText = 'Info Systems "Whereabouts" Manager'                    #--[ GUI Title ]--
$Script:EmailSubject = "Whereabouts"
$Script:SmtpServer = "mailserver.abc.org"
$Script:SmtpPort = "25"
$Script:RecipientEmailAddr = "GroupEmail@mycompany.com"                     #--[ Email to send to if below text is detected ]--
$Script:RecipientText = "Infotech"                                          #--[ Text to detect in recipient box ]--

#--[ Site List Options ]--------------------------------------------------------
$SiteList = @{    
    #--[ Format is "location, distance".  Distance is round trip miles.  Any distance of zero will not get written to log. ]--
    #--[ Add entries as needed.  Form size dynamically grows/shrinks according to this list. ]--
    "1" = @("Site 1","0");
    "2" = @("Site 2","0"); 
    "3" = @("Site 3","3");
    "4" = @("Alternate 4","2");
    "5" = @("Unknown","22");
    "6" = @("Test Site","0");
    "7" = @("Personal Business","0")    #--[ Remove the ; from the last entry ]--
    #"8" = @("x","6");
    #"9" = @("x","1");
    #"10" = @("x","0");
    #"11" = @("x","0");
    #"12" = @("Other","0")
}

#--[ Email Recipient Options ]--------------------------------------------------
$Script:Recipients = @()   #--[ List of recipients in case a group can't be used ]--
$Script:Recipients +="bob@abc.org"
If (!($Script:Debug)){     #--[ Use to block remaining recipients for test mode routing to sender ]--    
    $Script:Recipients +="sam@abc.org"
    $Script:Recipients +="karen@abc.org"
    $Script:Recipients +="robert@abc.org"
    $Script:Recipients +="nobody@gmail.com"
}

#--[ Functions ]--------------------------------------------------------------
Function UpdateData ($ItemNum,$SiteList){  #--[ Adds checkbox selected locations to the emqail and log file ]--
    $Script:Message += "- "+($SiteList["$ItemNum"])[0]+"<br>"
    If ((($SiteList["$ItemNum"])[1] -eq 0) -And ($Script:LogZero)){  #--[ Forces locations with a zero milage to get logged ]--
        $Script:RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }ElseIf (($SiteList["$ItemNum"])[1] -gt 0){
        $Script:RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }
}
Function ReloadForm {
	$Script:Form.Close()
	$Script:Form.Dispose()
    ActivateForm
    $Script:Stop = $True
}

Function KillForm {
	$Script:Form.Close()
    $Script:Form.Dispose()
    $Script:Stop = $True
}
Function UpdateOutput {  #--[ Refreshes the infobox contents ]--
    $Script:InfoBox.update()
    $Script:InfoBox.Select($InfoBox.Text.Length, 0)
    $Script:InfoBox.ScrollToCaret()
}

Function IsThereText ($TargetBox){  #--[ Checks for text in the text entry box(es) ]--
  if (($TargetBox.Text.Length -ge 8)){ 
    Return $true
  }else{
    Return $false
  }
}

Function SendEmail {    #--[ Constructs the email ]--
    $Smtp = New-Object Net.Mail.SmtpClient($Script:SmtpServer,$Script:SmtpPort)    
    #$Smtp.EnableSsl = $true
    #$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)    
    $Email = New-Object System.Net.Mail.MailMessage  
    $Signature = get-content -path "$Env:UserProfile\AppData\Roaming\Microsoft\Signatures\Default.txt"  #--[ Grabs default Outlook signature ]--
    $Script:Message += "<br>"
    Foreach ($Line in $Signature){
        $Script:Message += $Line+"<br>"
    } 
    $Script:Message += '</body></html>'
    $MessageBody = $Script:Message
    If ($Script:Console){write-host `nMessage Body:`n$MessageBody -ForegroundColor Magenta}
    $Email.IsBodyHTML = $true
    $Email.From = $Script:SenderEmail 
    If ($Script:RecipientTextBox.Text -eq $Script:RecipientText){  
        #$Script:RecipientEmail = $Script:RecipientEmail                  #--[ Assign a group email if you prefer ]--
        Foreach ($Person in $Script:Recipients){ 
            $Email.To.Add($Person)
        }
    }Else{
        $Script:RecipientEmailAddr = $Script:RecipientTextBox.Text 
    } 
    $Email.Subject = $Script:EmailSubject
    $ErrorActionPreference = "stop"
    $Email.Body = $MessageBody
    If ($Console){Write-Host `n"--- Email Sent ---" -ForegroundColor red }
    $Smtp.Send($Email)
}

Function LogTravel {      #--[ This is what gets written into the travel log file ]--
    If ($Script:Console){
        write-host `n"Adding the following to the log:" -ForegroundColor Yellow -NoNewline
        If ($Script:LogZero){
            Write-Host "`n   Zero mile entries will also be logged..." -ForegroundColor Green -NoNewline
        }
        Write-Host $Script:RunData -ForegroundColor Cyan
    } 
    $TravelLog = $Script:TravelLogPath+$Script:TravelLogFile
    If (!(Test-Path -Path $TravelLog)){
        New-Item -path $Script:TravelLogPath -name $Script:TravelLogFile -type "file" -value $Script:RunData
    }
    Try{
        Add-Content -path $TravelLog -value $Script:RunData -ErrorAction "stop"
    }Catch{
        If ($Script:Console){Write-Host "An error occurred writing to the log file..." -ForegroundColor Red}
    }
}

Function LockPC {  
    Invoke-Expression -Command "C:\Windows\System32\RUNDLL32.EXE user32.dll,LockWorkStation" -ErrorAction "stop"
}

Function LauncherIcon {       #--[ Create the Shortcut ]--
    $ErrorActionPreference = "stop"
    If (Test-Path "$Script:LaunchLocation\$Script:ScriptName.lnk"){
        If ($Script:Console){write-host "Found existing shortcut..." -ForegroundColor cyan}
    }Else{
        $WshShell = New-Object -ComObject WScript.Shell
        Try{
            If ($Script:Console){write-host "Creating new shortcut..." -ForegroundColor red}
            #$Shortcut = $WshShell.CreateShortcut("e:\temp3\$ScriptName.lnk")   #--[ Icon location for testing only ]--
            $Shortcut = $WshShell.CreateShortcut("$Script:LaunchLocation\$Script:ScriptName")
            $Shortcut.TargetPath = "C:\Windows\System32\windowspowershell\v1.0\powershell.exe"
            $Shortcut.IconLocation = "C:\Windows\System32\SHELL32.dll,156" #--[ icon index 156 is the mail envelope ]--
            $Shortcut.Arguments = "-windowstyle hidden -nonewwindow -Nop -Executionpolicy bypass -NoExit ""$PsScriptRoot\$Script:ScriptName.ps1"""
            $Shortcut.WorkingDirectory = $PSScriptRoot
            $Shortcut.Save()
        }Catch{
            $_.Exception.Message
        }
    }
}

Function ActivateForm {
    #--[ Create Form ]---------------------------------------------------------------------
    $Script:Form = New-Object System.Windows.Forms.Form    
    #$Script:Form.size = New-Object System.Drawing.Size($Script:FormWidth,$Script:FormHeight)
    #--[ The following allows the form to grow only downwards to expose the test mode checkbox ]--
    $Script:Form.minimumSize = New-Object System.Drawing.Size($Script:FormWidth,$Script:FormHeight)
    $Script:Form.maximumSize = New-Object System.Drawing.Size($Script:FormWidth,($Script:FormHeight+80))
    $Script:Notify = New-Object system.windows.forms.notifyicon
    $Script:Notify.icon = $Script:Icon              #--[ NOTE: Available tooltip icons are = warning, info, error, and none
    $Script:Notify.visible = $true
    [int]$Script:FormVTop = 0 
    [int]$Script:ButtonLeft = 55
    [int]$Script:ButtonTop = ($Script:FormHeight - 75)
    $Script:Form.Text = "$Script:ScriptName v$Script:ScriptVer"
    $Script:Form.StartPosition = "CenterScreen"
    $Script:Form.KeyPreview = $true
    $Script:Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Script:Form.Close();$Stop = $true}})
    $Script:ButtonFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",9,[System.Drawing.FontStyle]::Bold)

    #--[ Form Title Label ]-----------------------------------------------------------------
    $BoxLength = 350
    $LineLoc = 5
    $Script:FormLabelBox = new-object System.Windows.Forms.Label
    $Script:FormLabelBox.Font = $Script:ButtonFont
    $Script:FormLabelBox.Location = new-object System.Drawing.Size(($Script:FormHCenter-($BoxLength/2)-10),$LineLoc)
    $Script:FormLabelBox.size = new-object System.Drawing.Size($BoxLength,$Script:ButtonHeight)
    $Script:FormLabelBox.TextAlign = 2 
    $Script:FormLabelBox.Text = $TitleText
    $Script:Form.Controls.Add($Script:FormLabelBox)

    #--[ User Credential Label ]-------------------------------------------------------------
    $BoxLength = 250
    $LineLoc = 28
    $Script:UserCredLabel = New-Object System.Windows.Forms.Label 
    $Script:UserCredLabel.Location = New-Object System.Drawing.Point(($Script:FormHCenter-($BoxLength/2)-10),$LineLoc)
    $Script:UserCredLabel.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
    $Script:UserCredLabel.ForeColor = "DarkCyan"
    $Script:UserCredLabel.Font = $Script:ButtonFont
    $Script:UserCredLabel.Text = "Enter / Edit  Addressing Below:"
    $Script:UserCredLabel.TextAlign = 2 
    $Script:Form.Controls.Add($Script:UserCredLabel) 

    #--[ User ID Text Input Box ]-------------------------------------------------------------
    $BoxLength = 140
    $LineLoc = 55
    $Script:SenderTextBox = New-Object System.Windows.Forms.TextBox 
    $Script:SenderTextBox.Location = New-Object System.Drawing.Size(($Script:FormHCenter-158),$LineLoc)
    $Script:SenderTextBox.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
    $Script:SenderTextBox.TabIndex = 2

    If ($Script:SenderEmail -ne ""){
        $Script:SenderTextBox.Text = $Script:SenderEmail
        $Script:SenderTextBox.ForeColor = "Black"
    }Else{
        $Script:SenderTextBox.Text = "Sender's Email"
        $Script:SenderTextBox.ForeColor = "DarkGray"
    }
    $Script:SenderTextBox.TextAlign = 2
    $Script:SenderTextBox.Enabled = $True
    $Script:Form.Controls.Add($Script:SenderTextBox) 

    $Script:RecipientTextBox = New-Object System.Windows.Forms.TextBox 
    $Script:RecipientTextBox.Location = New-Object System.Drawing.Size((($Script:FormHCenter-3)),$LineLoc)
    $Script:RecipientTextBox.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
    $Script:RecipientTextBox.Text = $Script:DN
    $Script:RecipientTextBox.TabIndex = 3
    $Script:RecipientTextBox.ForeColor = "Black"
    If ($Script:RecipientEmail -eq $Script:RecipientEmailAddr){
        $Script:RecipientTextBox.Text = $Script:RecipientText
    }Else{
        $Script:RecipientTextBox.Text = $Script:RecipientEmail
    }    
    $Script:RecipientTextBox.TextAlign = 2
    $Script:RecipientTextBox.Enabled = $True
    $Script:RecipientTextBox.Add_LostFocus({
        if ($Script:RecipientTextBox.Text -eq '') {
            $Script:RecipientTextBox.Text = $Script:RecipientText
            $Script:RecipientEmail = $Script:RecipientEmailAddr
            $Script:RecipientTextBox.ForeColor = 'Black'
        }
    })
    $Script:Form.Controls.Add($Script:RecipientTextBox) 

    #--[ Information Box ]-------------------------------------------------------------------
    $BoxLength = 280
    $LineLoc = 90
    $Script:InfoBox = New-Object System.Windows.Forms.TextBox
    $Script:InfoBox.Location = New-Object System.Drawing.Size((($Script:FormHCenter-($BoxLength/2))-10),$LineLoc)
    $Script:InfoBox.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
    $Script:InfoBox.Text = "Select the location(s) you're going to:"
    $Script:InfoBox.Enabled = $False
    $Script:InfoBox.TextAlign = 2
    $Script:Form.Controls.Add($Script:InfoBox) #

    #--[ Option Buttons ]--------------------------------------------------------------------
    $CbLeft = 40
    $CbRight = 185
    $CbHeight = 120 #145
    $CbVar = 20
    $CbBox = 145

    $Count = 1
    While ($Count -le  $SiteList.Count){
       # $Count
        #--[ Left Checkbox ]--
        New-Variable -Name "CheckBox$Count" -Value $(new-object System.Windows.Forms.checkbox -Property @{
            Location = new-object System.Drawing.Size($CbLeft,$CbHeight)
            Size = new-object System.Drawing.Size($CbBox,$Script:TextHeight)
            Text = ($SiteList[$Count.ToString()])[0]
            Enabled = $true 
        })
        #Get-Variable -name "CheckBox$Count" -val
        $Form.Controls.Add($(Get-Variable -name "CheckBox$Count" -ValueOnly))  #
        #Remove-Variable -Name "CheckBox$Count"
        $Count++

        #--[ Right Checkbox ]--
        New-Variable -Name "CheckBox$Count" -Value $(new-object System.Windows.Forms.checkbox -Property @{
            Location = new-object System.Drawing.Size($CbRight,$CbHeight)
            Size = new-object System.Drawing.Size($CbBox,$Script:TextHeight)
            Text = ($SiteList[$Count.ToString()])[0]
            Enabled = $true 
        })
        $Form.Controls.Add($(Get-Variable -name "CheckBox$Count" -ValueOnly))  #
        #Remove-Variable -Name "CheckBox$Count"
        $Count++ 
        $CbHeight = $CbHeight+$CbVar
    }
    
#--[ Custom Message Box ]-------------------------------------------------------------------
$BoxLength = 280
#$LineLoc = $FormHeight-112
$LineLoc = $CbHeight+5 #$CbVar
$Script:CustomInfoBox = New-Object System.Windows.Forms.TextBox
$Script:CustomInfoBox.Location = New-Object System.Drawing.Size((($Script:FormHCenter-($BoxLength/2))-10),$LineLoc)
$Script:CustomInfoBox.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
$Script:CustomInfoBox.Text = "Enter any custom message here:"
$Script:CustomInfoBox.Enabled = $True
$Script:CustomInfoBox.TextAlign = 2
$Script:CustomInfoBox.Add_GotFocus({
    if ($Script:CustomInfoBox.Text -eq "Enter any custom message here:") {
        $Script:CustomInfoBox.Text = ''
        $Script:CustomInfoBox.ForeColor = 'Black'
    }
})
$Script:CustomInfoBox.Add_LostFocus({
    if ($Script:CustomInfoBox.Text -eq '') {
        $Script:CustomInfoBox.Text = "Enter any custom message here:"
        $Script:CustomInfoBox.ForeColor = 'Darkgray'
    }
})
$Script:Form.Controls.Add($Script:CustomInfoBox) #

#--[ Record Travel Checkbox ]--
$CbHeight = $LineLoc+28 #$CbVar
$TravelCheckBox = new-object System.Windows.Forms.checkbox
$TravelCheckBox.Location = new-object System.Drawing.Size($CbRight,$CbHeight)
$TravelCheckBox.Size = new-object System.Drawing.Size($CbBox,$Script:TextHeight)
$TravelCheckBox.Text = "Log Mileage ?"
$TravelCheckBox.Checked = $true #False
$TravelCheckBox.Enabled = $true 
$Form.Controls.Add($TravelCheckBox) 
    
#--[ Lock PC Checkbox ]--
$LockCheckBox = new-object System.Windows.Forms.checkbox
$LockCheckBox.Location = new-object System.Drawing.Size($CbLeft,$CbHeight)
$LockCheckBox.Size = new-object System.Drawing.Size($CbBox,$Script:TextHeight)
$LockCheckBox.Text = "Lock This PC ?"
$LockCheckBox.Checked = $true #False
$LockCheckBox.Enabled = $true 
$Form.Controls.Add($LockCheckBox) 

#--[ HIDDEN Tesing Checkbox & Message ]----------------------------------------------------
#--[ Drag the GUI window down to display this option for testing ]--
$BoxLength = 160
$LineLoc = $FormHeight-25
$TestCheckBox = new-object System.Windows.Forms.checkbox
$TestCheckBox.Location = new-object System.Drawing.Size(($Script:FormHCenter-($BoxLength/2)),$LineLoc)
$TestCheckBox.Size = new-object System.Drawing.Size($BoxLength,$Script:TextHeight)
$TestCheckBox.Text = "Execute in TEST Mode ?"
$TestCheckBox.Checked = $False
$TestCheckBox.Enabled = $true 
$TestCheckBox.Add_CheckStateChanged({
    if($TestCheckBox.checked){
        $LockCheckBox.Checked = $False   #--[ This forces the "locked" checkbox UNCHECKED when in test mode unless it's re-checked manually ]--
    }Else{
        $LockCheckBox.Checked = $true #False
    }
})
$Form.Controls.Add($TestCheckBox) 

$BoxLength = 280
$LineLoc = $FormHeight #$LineLoc-10
$Script:TestInfoBox = New-Object System.Windows.Forms.TextBox
$Script:TestInfoBox.Location = New-Object System.Drawing.Size((($Script:FormHCenter-($BoxLength/2))-10),$LineLoc)
$Script:TestInfoBox.Size = New-Object System.Drawing.Size($BoxLength,$Script:TextHeight) 
$Script:TestInfoBox.Text = "Test Mode forces the email to go to the sender only."
$Script:TestInfoBox.Enabled = $False
$Script:TestInfoBox.TextAlign = 2
$Script:Form.Controls.Add($Script:TestInfoBox) #

#--[ CLOSE Button ]------------------------------------------------------------------------
$BoxLength = 100
$LineLoc = $FormHeight-77
$Script:CloseButton = new-object System.Windows.Forms.Button
$Script:CloseButton.Location = New-Object System.Drawing.Size(($Script:FormHCenter-($BoxLength/2)-75),$LineLoc)
$Script:CloseButton.Size = new-object System.Drawing.Size($BoxLength,$Script:ButtonHeight)
$Script:CloseButton.TabIndex = 1
$Script:CloseButton.Text = "Cancel/Close"
$Script:CloseButton.Add_Click({
    #$Script:Form.close()
    #$Form.Close()
    #$Form.Dispose()
    KillForm
})
$Script:Form.Controls.Add($Script:CloseButton)

#--[ EXECUTE Button ]------------------------------------------------------------------------
$Script:ProcessButton = new-object System.Windows.Forms.Button
$Script:ProcessButton.Location = new-object System.Drawing.Size(($Script:FormHCenter-($BoxLength/2)+55),$LineLoc)
$Script:ProcessButton.Size = new-object System.Drawing.Size($BoxLength,$Script:ButtonHeight)
$Script:ProcessButton.Text = "Execute"
$Script:ProcessButton.Enabled = $True
$Script:ProcessButton.TabIndex = 5
$Script:ProcessButton.Add_Click({
    $Script:Message = '<!DOCTYPE html><html><head></head><body>'
    #--[ Testing Options ]---------------------------------------------
    If ($TestCheckBox.Checked){
        $LockCheckBox.Checked = $False
        $Script:Recipients = $Script:SenderEmail
        $Script:Message += "<font color=red><strong>--- RUNNING IN TEST MODE ---</strong></font><br><br>"
        #$Script:Message += "- "+($SiteList["$ItemNum"])[0]+"<br>"
        $Script:RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }
    #------------------------------------------------------------------
    $Script:Message += 'I am heading to the following location(s):<br>'

    $Count = 1
    While ($Count -le  $SiteList.Count){
        $IsChecked = (Get-Variable -name "CheckBox$Count" -ValueOnly)
        If ($IsChecked.Checked){
            UpdateData $Count $SiteList 
            $Flag++
        }
    $Count++
    }

    If ($CustomInfoBox -notlike "*custom message*"){
        $Script:Message += "<br>"+$CustomInfoBox.Text+"<br>"
        $Flag++
    }

    Add-Type -AssemblyName PresentationCore,PresentationFramework
    If ($Console){
        Write-host "Total location boxes checked ="$Flag -ForegroundColor Cyan
    }
    If ($Flag -eq 0){
        $ButtonType = [System.Windows.MessageBoxButton]::Ok
        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
        $MessageBody = "Whoa, dude, looks to me like you aren't actually going`n anywhere.  Perhaps you should select a destination?"
        $MessageTitle = "Huh?  Are you sure?"
        If ($Console){write-host `n"User forgot to select a destination... Recycling..." -ForegroundColor cyan}
        [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        $Script:Refresh = $True
        ReloadForm
    }Else{
        If ($TravelCheckBox.Checked){
            $Script:InfoBox.enabled = $true
            $Font = new-object System.Drawing.Font("Calibri",9,[System.Drawing.FontStyle]::Bold)
            #$Font = new-object System.Drawing.Font("Times New Roman",9,[System.Drawing.FontStyle]::Bold)
            $Script:InfoBox.Font = $font
            $Script:InfoBox.ForeColor = "yellow"
            $Script:InfoBox.BackColor = "green"            
            $Script:InfoBox.Text = "--- Logging Mileage ---"
            UpdateOutput
            Start-Sleep -sec 2
            $Script:InfoBox.Text = "-----------------------"
            UpdateOutput
            LogTravel
        }        
        Start-Sleep -sec 2
        $Script:InfoBox.Text = "--- Sending Email ---"
        UpdateOutput
        Start-Sleep -sec 2
        SendEmail $SmtpServer $SmtpPort
        $Script:InfoBox.Text = "-----------------------"
        UpdateOutput
        Start-Sleep -sec 2
        If ($LockCheckBox.Checked){
            $Script:InfoBox.Text = "--- Locking PC ---"
            UpdateOutput
            Start-Sleep -sec 2
            If (!($script:Debug)){
                LockPC 
            }
        }
    }
    $Script:Form.Close()
})
$Script:Form.Controls.Add($Script:ProcessButton)

#--[ Open Form ]---------------------------------------------------------------------------------
$Script:Form.topmost = $true
$Script:Form.Add_Shown({$Script:Form.Activate()})
[void] $Script:Form.ShowDialog()
}

#--[ End of Functions ]---------------------------------------------------------



LauncherIcon 

#--------------------------------[ Prep GUI ]----------------------------------- 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$Script:ScreenSize = (Get-WmiObject -Class Win32_DesktopMonitor | Select-Object ScreenWidth,ScreenHeight)
$Script:Width = $Script:ScreenSize.ScreenWidth
$Script:Height = $Script:ScreenSize.ScreenHeight

#--[ Define Form ]--------------------------------------------------------------
[int]$Script:FormWidth = 350
If ($SiteList.Count/2 -is [int]){
    [int]$Script:FormHeight = ((($SiteList.Count/2)*20)+255)   #--[ Dynamically Created Variable for Box Size (Even count) ]--
}Else{
    [int]$Script:FormHeight = ((($SiteList.Count/2)*23)+255)   #--[ Dynamically Created Variable for Box Size (Odd count) ]--
}
If ($Debug){
    Write-Host "Form Height = "$Script:FormHeight
    Write-Host "Form Width  = "$Script:FormWidth
    Write-Host "Site Count  = "$SiteList.Count
    Write-Host "Box Count   = "($SiteList.Count/2)    
}

[int]$Script:FormHCenter = ($Script:FormWidth / 2)   # 170 Horizontal center point
[int]$Script:FormVCenter = ($Script:FormHeight / 2)  # 209 Vertical center point
[int]$Script:ButtonHeight = 25
[int]$Script:TextHeight = 20

ActivateForm 

if($Script:Stop -eq $true){$Script:Form.Close();break;break}
If ($Console){write-host $Script:RecipientTextBox.Text -ForegroundColor cyan}

<#--[ Manual Shortcut Details ]---------------------------------------- 
To prevent any pop-up commend windows use the following in the "Target" field of a shortcut 
 
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -file "c:\scripts\Whereabouts.ps1" -windowstyle hidden -nonewwindow 
 
- Adjust the path to the script as needed. 
- Set the "Run" option to "Minimized" 
- An icon will appear briefly in the taskbar while assemblies load, then disappear as the GUI loads. 
 
#>