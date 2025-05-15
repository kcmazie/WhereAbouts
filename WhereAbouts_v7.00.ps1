<#------------------------------------------------------------------------------ 
         File Name : Whereabouts.ps1 
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com) 
                   : 
       Description : Programatically creates an email to send to the predetermined user or group for notification about 
                   : where you are.   Grabs Outlook email signature from current users profile.  Determines sender by 
                   : current logged on user. 
                   : 
             Notes : Normal operation is with no command line options.  See end of script for detail about how to 
                   : launch via shortcut.  Edit the external XML file to include required run variables.  An example 
                   : XML file is included at the end of the script.  The expectation is that you would normally
                   : send emails to a group email or list of emails.  Both of these are pre-programmed into to 
                   : external XML file.  If these are left out then the manually entered email in the GUI will be
                   : used.  If the manul email is left off or not formatted the script fails with a warning.  The size 
                   : of the GUI dynamically adjusts according to the list of sites in the XML file.  An option exists 
                   : to add an optional message to the outgoing email.  Check boxes exist to enable/disable automatic
                   : logging of mileage to sites and to lock the local PC.  Additionally the GUI can be expanded down
                   : to expose a testing option box that alters operation forcing the email to only go to the sender.
                   : Test mode also alters the mileage log file name so as to not over-write the live file.  If executed
                   : from an editor such as VS Code debugging messages are automatically displayed.  An option is also
                   : included to check for and create a task bar shortcut to run the script if the option is enabled
                   : in the XML file.  Sender's email is automatically determined to be the logged-in user.
                   : 
         Arguments : Command line options for testing:  None
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
                   : v5.20 - 03-30-22 - Added 0 mile logging, reordered variable location, adjusted some console 
                   :                    message text messages for clarity                
                   : v6.00 - 05-05-23 - Relocated options out to XML file for publishing.  Refactored some sections.
                   : v6.10 - 10-02-23 - Switched out WMI for CIM
                   : v7.00 - 05-14-25 - Major rewrite.  Added extensive use of external config file for run options.
                   :                    Added more debug messages.  Removed some items hardcoded in previous 
                   :                    releases.  Added detection of run from editor to force debug mode.  Added detection
                   :                    of Html verses text based email sigantures.
                   : #>
                   $ScriptVer = "7.00"    <#--[ Current version # used in script ]--
                   : 
------------------------------------------------------------------------------#>
#Requires -Version 5.1
Clear-Host 

#--[ For Testing ]-------------
#$Debug = $true
#------------------------------

#--[ Suppress Console ]-------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition ' 
[DllImport("Kernel32.dll")] 
public static extern IntPtr GetConsoleWindow(); 
 
[DllImport("user32.dll")] 
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow); 
' 
$ConsolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($ConsolePtr, 0) | Out-Null
#------------------------------------------------------------------------------#
 
#--[ Runtime Variables ]----------------------------------------------------
$ErrorActionPreference = "silentlycontinue"
$Icon = [System.Drawing.SystemIcons]::Information
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0] 
$Message = ""
$RunData = ""

#--[ Functions ]--------------------------------------------------------------

Function GetConsoleHost ($ExtOption){  #--[ Detect if we are using a script editor or the console ]--
    Switch ($Host.Name){
        'consolehost'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $False -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell Console detected." -Force
        }
        'Windows PowerShell ISE Host'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell ISE editor detected." -Force
        }
        'PrimalScriptHostImplementation'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "COnsoleMessage" -Value "PrimalScript or PowerShell Studio editor detected." -Force
        }
        "Visual Studio Code Host" {
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "Visual Studio Code editor detected." -Force
        }
    }
    If ($ExtOption.ConsoleState){
        StatusMsg "Detected session running from an editor..." "Magenta" $ExtOption
    }
    Return $ExtOption
}

Function LoadConfig ($ExtOption,$ConfigFile,$Debug){  #--[ Read and load configuration file ]-------------------------------------
    if (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
        If ($Debug){
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $True 
        }Else{
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $False
        }
        StatusMsg "Loading external config file..." "Magenta" $ExtOption
        [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--    
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ShortcutLocation" -Value "$Env:AppData\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Shortcut" -Value $Config.Settings.General.Shortcut
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TravelLogFile" -Value $Config.Settings.General.TravelLogFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TravelLogPath" -Value "$Env:UserProfile\Documents\"
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ZeroLog" -Value $Config.Settings.General.ZeroLog
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TitleText" -Value $Config.Settings.General.TitleText
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.Email.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpPort" -Value $Config.Settings.Email.SmtpPort
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SSL" -Value $Config.Settings.Email.SSL
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailUser" -Value $Config.Settings.Email.EmailUser
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailPass" -Value $Config.Settings.Email.EmailPass
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSubject" -Value $Config.Settings.Email.EmailSubject
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PresetGroupEmail" -Value $Config.Settings.Email.PresetGroupEmail
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PresetGroupText" -Value $Config.Settings.Email.PresetGroupText

        If ((Get-Date).DayOfWeek -eq $Config.Settings.General.DayOfWeek){  #--[ Triggers email to group on selected day of week ]--
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Alert" -Value $True
        }Else{
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Alert" -Value $False
        }
        
        #--[ Site List Options.  Site list is located within the external XML file.  ]-------------------------------------------  
        #--[ Format is "location, distance".  Distance is round trip miles.  Any distance of zero will not get written to log ]--
        #--[ unless logzero option is enabled.  Add entries to XML as needed.  Form size dynamically grows/shrinks according  ]--
        #--[ to that list. ]--
        $SiteList = [Ordered]@{}  
        $Index = 1
        ForEach($Site in $Config.Settings.Sites.site){
            $SiteList.Add($Index, @($Site))
            $Index++
        }
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SiteList" -Value $SiteList

        #--[ Email Recipient Options ]--------------------------------------------------
        $Recipients = @()   #--[ List of recipients in case a group can't be used ]--
        ForEach($Recipient in $Config.Settings.Email.Recipients.Recipient){
            $Recipients +=$Recipient
        }
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Recipients" -Value $Recipients

        #--[ Email Options ]------------------------------------------------------------
        $SenderEmail = $Env:Username+"@"+($ENV:UserDNSdomain.ToLower())
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SenderEmail" -Value $SenderEmail
        #--[ Test for email siganture ]--
        If (Test-Path -path "$Env:UserProfile\AppData\Roaming\Microsoft\Signatures" -pathtype Container){
            $List = Get-ChildItem  -path ($Env:UserProfile+'\AppData\Roaming\Microsoft\Signatures')
            ForEach ($File in $List ){ 
                If ($File -like "*.htm*"){
                    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSig" -Value "True"
                    Break
                }ElseIf($File -like "*.txt*"){
                    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSig" -Value "True"
                    Break
                }Else{
                    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSig" -Value "False"
                }
            }
            $Sig = get-content -path ($Env:UserProfile+'\AppData\Roaming\Microsoft\Signatures\'+$File)
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSignature" -Value $Sig
        }Else{
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSig" -Value "False"
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSignature" -Value ""
        }
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $ExtOption
        break;break;break
    }
    Return $ExtOption
}

Function SendEmail ($ExtOption){    #--[ Constructs the email ]--
    $Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,$ExtOption.SmtpPort)  
    If ($ExtOption.SSL -like "*True*"){  
        $Smtp.EnableSsl = $True
        $Smtp.Credentials = New-Object System.Net.NetworkCredential($ExtOption.EmailUser,$ExtOption.EmailPass)    
    }
    $Email = New-Object System.Net.Mail.MailMessage  
    $Email.Subject = $ExtOption.EmailSubject
    If ($ExtOption.Email.HTML -like "*true*"){
        $Email.IsBodyHTML = $true
    }Else{
        $Email.IsBodyHTML = $false
    }
    $Email.From = $ExtOption.SenderEmail 
    Foreach ($Person in $ExtOption.Recipients){ 
        $Email.To.Add($Person)
    }
    $ErrorActionPreference = "stop"
    $Email.Body = $ExtOption.EmailMessage
    StatusMsg ("Sending email FROM : "+$Email.From) "Yellow" $ExtOption
    StatusMsg ("Sending emails TO  : "+$Email.To) "Yellow" $ExtOption

    Try{
        $Smtp.Send($Email)
        StatusMsg " --- Email Sent ---" "Green" $ExtOption
    }Catch{
        StatusMsg " --- Email send has FAILED ---" "Red" $ExtOption
    }
}

Function UpdateData ($ItemNum,$SiteList){  #--[ Adds checkbox selected locations to the email and log file ]--
    $Message += "- "+($SiteList["$ItemNum"])[0]+"<br>"
    If ((($SiteList["$ItemNum"])[1] -eq 0) -And ($ExtOption.ZeroLog)){  #--[ Forces locations with a zero milage to get logged ]--
        $RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }ElseIf (($SiteList["$ItemNum"])[1] -gt 0){
        $RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }
}#>

Function ReloadForm ($ExtOption){
    StatusMsg "Reloading WinForm..." "Yellow" $ExtOption
    $Form.Close()
    $Form.Dispose()
    ActivateForm $ExtOption
    $Stop = $True
}

Function KillForm {
    $Form.Close()
    $Form.Dispose()
    $Stop = $True
}
Function UpdateOutput {  #--[ Refreshes the infobox contents ]--
    $InfoBox.update()
    $InfoBox.Select($InfoBox.Text.Length, 0)
    $InfoBox.ScrollToCaret()
}

Function IsThereText ($TargetBox){  #--[ Checks for text in the text entry box(es) ]--
  if (($TargetBox.Text.Length -ge 8)){ 
    Return $true
  }else{
    Return $false
  }
}

Function LogTravel ($ExtOption, $RunData){      #--[ This is what gets written into the travel log file ]--
    $ErrorActionPreference = "stop"    
    If ($ExtOption.LogZero){
        StatusMsg "Zero mile entries will also be logged..." "Green" $ExtOption
    }
    $TravelLog = $ExtOption.TravelLogPath+$ExtOption.TravelLogFile
    If (!(Test-Path -Path $TravelLog)){
        New-Item -path $TravelLog -type "file" 
        StatusMsg "Creating new travel log file..." "Magenta" $ExtOption
    }
    Try{
        Add-Content -path $TravelLog -value $RunData -ErrorAction "stop"
        StatusMsg "Adding location data to the log..." "magenta" $ExtOption
    }Catch{
        $_.Exception.Message
        StatusMsg "An error occurred writing to the log file..." "Red" $ExtOption
    }
}

Function LockPC {  
    StatusMsg "Locking PC session..." "Red" $ExtOption
    Invoke-Expression -Command "C:\Windows\System32\RUNDLL32.EXE user32.dll,LockWorkStation" -ErrorAction "stop"
}

Function StatusMsg ($Msg, $Color, $ExtOption){
    $ErrorActionPreference = "SilentlyContinue"
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    If ($ExtOption.Debug -or $ExtOption.ConsoleState){
        Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    }
    $Msg = ""
}

Function PopupMessage ($MessageBody,$MessageTitle,$DebugMsg){
    $ButtonType = [System.Windows.MessageBoxButton]::Ok
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
    #$MessageBody = "You must enter a valid email address or this script will crash.  The script will now recycle...?"
    #$MessageTitle = "Invalid Email Address !!!"
    StatusMsg $DebugMsg "red" $ExtOption
    [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    ReloadForm $ExtOption
}

Function ActivateForm ($ExtOption){
    #--[ Create Form ]---------------------------------------------------------------------
    $Form = New-Object System.Windows.Forms.Form    
    #--[ The following allows the form to grow only downwards to expose the test mode checkbox ]--
    $Form.minimumSize = New-Object System.Drawing.Size($ExtOption.FormWidth,$ExtOption.FormHeight)
    $Form.maximumSize = New-Object System.Drawing.Size($ExtOption.FormWidth,($ExtOption.FormHeight+80))
    $Notify = New-Object system.windows.forms.notifyicon
    $Notify.icon = $Icon              #--[ NOTE: Available tooltip icons are = warning, info, error, and none
    $Notify.visible = $true
    [int]$FormVTop = 0 
    [int]$ButtonLeft = 55
    [int]$ButtonTop = ($ExtOption.FormHeight - 75)
    $Form.Text = "$ScriptName v$ScriptVer"
    $Form.StartPosition = "CenterScreen"
    $Form.KeyPreview = $true
    $Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Form.Close();$Stop = $true}})
    
    $LabelFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",10,[System.Drawing.FontStyle]::Bold)
    $ButtonFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",10,[System.Drawing.FontStyle]::Regular)
    $TextBoxFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",9,[System.Drawing.FontStyle]::Regular)
    $CheckBoxFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",9,[System.Drawing.FontStyle]::Regular)
 
    #--[ Form Title Label ]-----------------------------------------------------------------
    $BoxLength = 350
    $LineLoc = 5
    $FormLabelBox = new-object System.Windows.Forms.Label
    $FormLabelBox.Font = $LabelFont
    $FormLabelBox.Location = new-object System.Drawing.Size(($ExtOption.FormHCenter-($BoxLength/2)-10),$LineLoc)
    $FormLabelBox.size = new-object System.Drawing.Size($BoxLength,$ExtOption.ButtonHeight)
    $FormLabelBox.TextAlign = 2 
    $FormLabelBox.Text = $ExtOption.TitleText
    $Form.Controls.Add($FormLabelBox)

    #--[ User Credential Label ]-------------------------------------------------------------
    $BoxLength = 250
    $LineLoc = 28
    $UserCredLabel = New-Object System.Windows.Forms.Label 
    $UserCredLabel.Location = New-Object System.Drawing.Point(($ExtOption.FormHCenter-($BoxLength/2)-10),$LineLoc)
    $UserCredLabel.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $UserCredLabel.ForeColor = "DarkCyan"
    $UserCredLabel.Font = $TextBoxFont
    $UserCredLabel.Text = "Enter / Edit  Addressing Below:"
    $UserCredLabel.TextAlign = 2 
    $Form.Controls.Add($UserCredLabel) 

    #--[ User ID Text Input Box ]-------------------------------------------------------------
    $BoxLength = 140
    $LineLoc = 55
    $SenderTextBox = New-Object System.Windows.Forms.TextBox 
    $SenderTextBox.Location = New-Object System.Drawing.Size(($ExtOption.FormHCenter-158),$LineLoc)
    $SenderTextBox.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $SenderTextBox.Font = $TextBoxFont
    $SenderTextBox.TabIndex = 2

    If ($SenderEmail -ne ""){
        $SenderTextBox.Text = $ExtOption.SenderEmail
        $SenderTextBox.ForeColor = "Black"
    }Else{
        $SenderTextBox.Text = "Sender's Email"
        $SenderTextBox.ForeColor = "DarkGray"
    }
    $SenderTextBox.TextAlign = 2
    $SenderTextBox.Enabled = $True
    $Form.Controls.Add($SenderTextBox) 

    #--[ Recipient Text Box ]-------------------------------------
    $RecipientTextBox = New-Object System.Windows.Forms.TextBox 
    $RecipientTextBox.Location = New-Object System.Drawing.Size((($ExtOption.FormHCenter-3)),$LineLoc)
    $RecipientTextBox.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $RecipientTextBox.Text = $ExtOption.GroupEmail
    $RecipientTextBox.Font = $TextBoxFont
    $RecipientTextBox.TabIndex = 3
    $RecipientTextBox.ForeColor = "Black"
    $EmailFlag = $False
    If ($Null -ne $ExtOption.PresetGroupEmail){  #--[ If preset group email address exists, use it ]--
        $RecipientTextBox.Text = $ExtOption.PresetGroupText
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Recipients" -Value $ExtOption.PresetGroupEmail
        $EmailFlag = $True
    }ElseIf ($Null -ne $ExtOption.Recipients[0]){  #--[ Else use preset recipient list if present (pre-loaded from XML) ]--
        $RecipientTextBox.Text = "Using Preset List"
        $EmailFlag = $True
    }Else{   #--[ Else expect a manually entered recipient ]--     
        $RecipientTextBox.Text = "Enter Email Recipient Here" #$RecipientEmail
    }    
    $RecipientTextBox.TextAlign = 2
    $RecipientTextBox.Enabled = $True
    $Form.Controls.Add($RecipientTextBox) 

    #--[ Information Box ]-------------------------------------------------------------------
    $BoxLength = 280
    $LineLoc = 90
    $InfoBox = New-Object System.Windows.Forms.TextBox
    $InfoBox.Location = New-Object System.Drawing.Size((($ExtOption.FormHCenter-($BoxLength/2))-10),$LineLoc)
    $InfoBox.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $InfoBox.Text = "Select the location(s) you're going to:"
    $InfoBox.Font = $TextBoxFont
    $InfoBox.Enabled = $False
    $InfoBox.TextAlign = 2
    $Form.Controls.Add($InfoBox) #

    #--[ Option Buttons ]--------------------------------------------------------------------
    $CbLeft = 40
    $CbRight = 185
    $CbHeight = 120 #145
    $CbVar = 20
    $CbBox = 145
   
    #--[ Dynamically grow the box size depending on number of entries in XML ]--
    $Count = 0

    While (($ExtOption.SiteList).Count -gt $Count) {
        #--[ Left Checkbox ]--
        Remove-Variable -Name "CheckBox$Count" -ErrorAction SilentlyContinue
        $Left = new-object System.Windows.Forms.checkbox -Property @{
            Location = new-object System.Drawing.Size($CbLeft,$CbHeight)
            Size = new-object System.Drawing.Size($CbBox,$ExtOption.TextHeight)
            Font = $CheckBoxFont
            Text = ($ExtOption.SiteList[$Count])[0].Split(",")[0]
            Enabled = $true 
        }
        New-Variable -Name "CheckBox$Count" -value $Left
        $LeftBox = Get-Variable -name "CheckBox$Count" -ValueOnly
        $Form.Controls.Add($LeftBox) 
        $Count++
        #--[ Right Checkbox ]--
        Remove-Variable -Name "CheckBox$Count" -ErrorAction SilentlyContinue
        $Right = New-Object System.Windows.Forms.checkbox -Property @{
            Location = new-object System.Drawing.Size($CbRight,$CbHeight)
            Size = new-object System.Drawing.Size($CbBox,$ExtOption.TextHeight)
            Text = ($ExtOption.SiteList[$Count])[0].Split(",")[0]
            Font = $CheckBoxFont
            Enabled = $true 
        }
        New-Variable -Name "CheckBox$Count" -value $Right
        $RightBox = Get-Variable -name "CheckBox$Count" -ValueOnly
        $Form.Controls.Add($RightBox)  
        $Count++ 
     
        $CbHeight = $CbHeight+$CbVar
    }
    
    #--[ Custom Message Box ]-------------------------------------------------------------------
    $BoxLength = 280
    #$LineLoc = $ExtOption.FormHeight-112
    $LineLoc = $CbHeight+5 #$CbVar
    $CustomInfoBox = New-Object System.Windows.Forms.TextBox
    $CustomInfoBox.Location = New-Object System.Drawing.Size((($ExtOption.FormHCenter-($BoxLength/2))-10),$LineLoc)
    $CustomInfoBox.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $CustomInfoBox.Text = "Enter any custom message here:"
    $CustomInfoBox.Font = $TextBoxFont
    $CustomInfoBox.Enabled = $True
    $CustomInfoBox.TextAlign = 2
    $CustomInfoBox.Add_GotFocus({
        if ($CustomInfoBox.Text -eq "Enter any custom message here:") {
            $CustomInfoBox.Text = ''
            $CustomInfoBox.ForeColor = 'Black'
        }
    })
    $CustomInfoBox.Add_LostFocus({
        if ($CustomInfoBox.Text -eq '') {
            $CustomInfoBox.Text = "Enter any custom message here:"
            $CustomInfoBox.ForeColor = 'Darkgray'
        }
    })
    $Form.Controls.Add($CustomInfoBox) #

    #--[ Record Travel Checkbox ]--
    $CbHeight = $LineLoc+28 #$CbVar
    $TravelCheckBox = new-object System.Windows.Forms.checkbox
    $TravelCheckBox.Location = new-object System.Drawing.Size($CbRight,$CbHeight)
    $TravelCheckBox.Size = new-object System.Drawing.Size($CbBox,$ExtOption.TextHeight)
    $TravelCheckBox.Font = $CheckBoxFont
    $TravelCheckBox.Text = "Log Mileage ?"
    $TravelCheckBox.Checked = $true #False
    $TravelCheckBox.Enabled = $true 
    $Form.Controls.Add($TravelCheckBox) 
    
    #--[ Lock PC Checkbox ]--
    $LockCheckBox = new-object System.Windows.Forms.checkbox
    $LockCheckBox.Location = new-object System.Drawing.Size($CbLeft,$CbHeight)
    $LockCheckBox.Size = new-object System.Drawing.Size($CbBox,$ExtOption.TextHeight)
    $LockCheckBox.Font = $CheckBoxFont
    $LockCheckBox.Text = "Lock This PC ?"
    $LockCheckBox.Checked = $true #False
    $LockCheckBox.Enabled = $true 
    $Form.Controls.Add($LockCheckBox) 

    #--[ HIDDEN Tesing Checkbox & Message ]----------------------------------------------------
    #--[ Drag the GUI window down to display this option for testing ]--
    $BoxLength = 190
    $LineLoc = $ExtOption.FormHeight-25
    $TestCheckBox = new-object System.Windows.Forms.checkbox
    $TestCheckBox.Location = new-object System.Drawing.Size(($ExtOption.FormHCenter-($BoxLength/2)),$LineLoc)
    $TestCheckBox.Size = new-object System.Drawing.Size($BoxLength,$ExtOption.TextHeight)
    $TestCheckBox.Font = $CheckBoxFont
    $TestCheckBox.Text = "Execute in TEST Mode ?"
    $TestCheckBox.Checked = $False
    $TestCheckBox.Enabled = $true 
    $TestCheckBox.Add_CheckStateChanged({
        if($TestCheckBox.checked){
            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TravelLogFile" -Value "Testlog.csv"
            $LockCheckBox.Checked = $False   #--[ This forces the "locked" checkbox UNCHECKED when in test mode unless it's re-checked manually ]--
        }Else{
            $LockCheckBox.Checked = $true 
        }
    })
    $Form.Controls.Add($TestCheckBox) 

    $BoxLength = 280
    $LineLoc = $ExtOption.FormHeight #$LineLoc-10
    $TestInfoBox = New-Object System.Windows.Forms.TextBox
    $TestInfoBox.Location = New-Object System.Drawing.Size((($ExtOption.FormHCenter-($BoxLength/2))-10),$LineLoc)
    $TestInfoBox.Size = New-Object System.Drawing.Size($BoxLength,$ExtOption.TextHeight) 
    $TestInfoBox.Text = "Test Mode forces the email to go to the sender only."
    $TestCheckBox.Font = $TextBoxFont
    $TestInfoBox.Enabled = $False
    $TestInfoBox.TextAlign = 2
    $Form.Controls.Add($TestInfoBox) #

    #--[ CLOSE Button ]------------------------------------------------------------------------
    $BoxLength = 100
    $LineLoc = $ExtOption.FormHeight-90
    $CloseButton = new-object System.Windows.Forms.Button
    $CloseButton.Location = New-Object System.Drawing.Size(($ExtOption.FormHCenter-($BoxLength/2)-75),$LineLoc)
    $CloseButton.Size = new-object System.Drawing.Size($BoxLength,$ExtOption.ButtonHeight)
    $CloseButton.Font = $ButtonFont
    $CloseButton.TabIndex = 1
    $CloseButton.Text = "Cancel/Close"
    $CloseButton.Add_Click({
        KillForm
    })
    $Form.Controls.Add($CloseButton)

    #--[ EXECUTE Button ]------------------------------------------------------------------------
    $ProcessButton = new-object System.Windows.Forms.Button
    $ProcessButton.Location = new-object System.Drawing.Size(($ExtOption.FormHCenter-($BoxLength/2)+55),$LineLoc)
    $ProcessButton.Size = new-object System.Drawing.Size($BoxLength,$ExtOption.ButtonHeight)
    $ProcessButton.Font = $ButtonFont
    $ProcessButton.Text = "Execute"
    $ProcessButton.Enabled = $True
    $ProcessButton.TabIndex = 5
    $ProcessButton.Add_Click({
        $ErrorActionPreference = "silentlycontinue" 
        #--[ Testing Options ]---------------------------------------------
        If ($TestCheckBox.Checked){
            $LockCheckBox.Checked = $False
            $ExtOption | Add-Member -MemberType NoteProperty -Name "Recipients" -Value $ExtOption.SenderEmail -force
            $Message = "<font color=red><strong>--- RUNNING IN TEST MODE ---</strong></font><br><br>"
        }Else{
            $Message = '<!DOCTYPE html><html><head></head><body>'
        }
        #------------------------------------------------------------------
        $Message += 'I am heading to the following location(s):<br>'
        $Counter = 0
        $Count = 0
        Statusmsg ('Total locations in list: '+$ExtOption.SiteList.Count.ToString()) "Magenta" $ExtOption
        While ($ExtOption.SiteList.Count -gt  $Counter){
            $IsChecked = Get-Variable -name ("*CheckBox$Counter") -ValueOnly
            If ($IsChecked.CheckState -eq "Checked"){
                $Count++
                $Selected = $ExtOption.SiteList[$Counter]
                $Msg = $Selected.Split(",")[0]+" Selected.  ("+$Selected.Split(",")[1]+" Miles)"
                Statusmsg $Msg "Cyan" $ExtOption
                $Message += " - "+($Selected.Split(","))[0]+"<br>"   #--[ Adds checkbox selected locations to the email ]--
                If (($Selected.Split(",")[1] -eq 0) -And ($ExtOption.LogZero)){  #--[ Forces locations with a zero milage to get logged ]--
                    $RunData += "`n"+(Get-Date).toShortDateString()+","+$Selected.Split(",")[0]+","+$Selected.Split(",")[1]
                }ElseIf ($Selected.Split(",")[1] -gt 0){
                    $RunData += "`n"+(Get-Date).toShortDateString()+","+$Selected.Split(",")[0]+","+$Selected.Split(",")[1]
                }
            }
            $Counter++
        }

        If ($CustomInfoBox -notlike "*custom message*"){
            $Message += "<br>"+$CustomInfoBox.Text+"<br>"
        }
 
        Add-Type -AssemblyName PresentationCore,PresentationFramework 
        StatusMsg ("Total location selected: "+$Count) "Magenta" $ExtOption
    
        If ($Count -eq 0){
            $MsgBody = "Whoa, dude, looks to me like you aren't actually going`n anywhere.  Perhaps you should select a destination?"
            $MsgTitle = "Huh?  Are you sure?"
            $MsgDebug = "User forgot to select a destination... Recycling..."
            PopupMessage $MsgBody $MsgTitle $MsgDebug
        }Else{
            If ($RecipientTextBox.Text -like "*@*"){  #--[ Email will send to text entered in form IF is a valid addr  ]--
                $ExtOption | Add-Member -MemberType NoteProperty -Name "Recipients" -Value $RecipientTextBox.Text -force 
                StatusMsge "No preset recipients found.  Using manually entered recipient(s)..." "Magenta" $ExtOption
            }ElseIf (($EmailFlag) -and ($RecipientTextBox.Text -notlike "*@*")){
                StatusMsg "Using preset recipient email(s) from XML..." "Magenta" $ExtOption
            }Else{
                $MsgBody = "No destination email address was found. You must enter a valid email address or this script will crash.  You must manually enter a destination email address or preset one or more addresses within the external option file.  The script will now restart..."
                $MsgTitle = "Invalid Email Address !!!"
                $MsgDebug = "User did not enter a valid email address... Recycling..."
                PopupMessage $MsgBody $MsgTitle $MsgDebug
            }
            If ($TravelCheckBox.Checked){
                $InfoBox.enabled = $true
                $Font = new-object System.Drawing.Font("Calibri",9,[System.Drawing.FontStyle]::Bold)
                #$Font = new-object System.Drawing.Font("Times New Roman",9,[System.Drawing.FontStyle]::Bold)
                $InfoBox.Font = $font
                $InfoBox.ForeColor = "yellow"
                $InfoBox.BackColor = "green"            
                $InfoBox.Text = "--- Logging Mileage ---"
                UpdateOutput  #--[ Refresh the infobox ]--
                Start-Sleep -sec 2
                $InfoBox.Text = "-----------------------"
                UpdateOutput  #--[ Refresh the infobox ]--
                LogTravel $ExtOption $Rundata
            }   
            If ($ExtOption.EmailSig -eq "False"){
                StatusMsg "Text email signature detected..." "Magenta" $ExtOption
                $Message += '</body></html>'
                $NewSig = $Message
            }Else{
                $Message += '<br>'
                StatusMsg "HTML email signature detected..." "Magenta" $ExtOption
                StatusMsg "Modifying HTML email signature on the fly..." "Magenta" $ExtOption
                ForEach ($line in $ExtOption.EmailSignature){
                    $NewSig += $line
                    if ($line -eq "<div class=WordSection1>"){
                        $newsig += $Message
                    }
                }
            }

            $ExtOption | Add-Member -MemberType NoteProperty -Name "EmailMessage" -Value $NewSig -force

            Start-Sleep -sec 2
            $InfoBox.Text = "--- Sending Email ---"
            UpdateOutput  #--[ Refresh the infobox ]--
            Start-Sleep -sec 2
            SendEmail $ExtOption
            $InfoBox.Text = "-----------------------"
            UpdateOutput  #--[ Refresh the infobox ]--
            Start-Sleep -sec 2
            If ($LockCheckBox.Checked){
                $InfoBox.Text = "--- Locking PC ---"
                UpdateOutput  #--[ Refresh the infobox ]--
                Start-Sleep -sec 2
                #If (!($Script:Debug)){
 #                   LockPC $ExtOption
                #}
            }
        }
        $Form.Close()
    })
    $Form.Controls.Add($ProcessButton)

    #--[ Open Form ]---------------------------------------------------------------------------------
    $Form.topmost = $true
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
    if($Stop -eq $true){$Form.Close();break;break}
    Return
}

#--[ End of Functions ]---------------------------------------------------------

#--[ Load external XML options file ]--
$ExtOption = New-Object -TypeName psobject #--[ Object to hold runtime options ]--
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
$ExtOption = GetConsoleHost $ExtOption #--[ Detect Runspace, add to option file ]--

StatusMsg "Starting run..." "Yellow" $ExtOption
If ($ExtOption.ConsoleState){ 
    StatusMsg $ExtOption.ConsoleMessage "Cyan" $ExtOption
}
$ExtOption = LoadConfig $ExtOption $ConfigFile $Debug

If ($ExtOption.Shortcut -eq "True"){  #--[ Create the Shortcut ]--
    $ErrorActionPreference = "stop"
    If (Test-Path "$ShortcutLocation\$ScriptName.lnk"){
        StatusMsg "Found existing shortcut..." "Magenta" $ExtOption
    }Else{
        $WshShell = New-Object -ComObject WScript.Shell
        Try{
            StatusMsg "Creating new shortcut..." "red" $ExtOption
            #$Shortcut = $WshShell.CreateShortcut("e:\temp3\$ScriptName.lnk")   #--[ Icon location for testing only ]--
            $Shortcut = $WshShell.CreateShortcut("$LaunchLocation\$ScriptName.lnk")
            $Shortcut.TargetPath = "C:\Windows\System32\windowspowershell\v1.0\powershell.exe"
            $Shortcut.IconLocation = "C:\Windows\System32\SHELL32.dll,156" #--[ icon index 156 is the mail envelope ]--
            $Shortcut.Arguments = "-windowstyle hidden -nonewwindow -Nop -Executionpolicy bypass -NoExit ""$PsScriptRoot\$ScriptName.ps1"""
            $Shortcut.WorkingDirectory = $PSScriptRoot
            $Shortcut.Save()
        }Catch{
            $_.Exception.Message
        }
    }
}

#--[ Prep GUI ]------------------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$ScreenSize = (Get-CimInstance -Class Win32_DesktopMonitor | Select-Object ScreenWidth,ScreenHeight)

If ($ScreenSize.Count -gt 1){  #--[ Detect multiple monitors ]--
    StatusMsg "More than 1 monitor detected..." "Magenta"
    ForEach ($Resolution in $ScreenSize){
        If ($Null -ne $Resolution.ScreenWidth){
            $ScreenWidth = $Resolution.ScreenWidth
            $ScreenHeight = $Resolution.ScreenHeight
            Break
        }
    }
}Else{
    $ScreenWidth = $Resolution.ScreenWidth
    $ScreenHeight = $Resolution.ScreenHeight
}
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ScreenWidth" -Value $ScreenWidth
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ScreenHeight" -Value $ScreenHeight

#--[ Define Form ]--------------------------------------------------------------
[int]$FormWidth = 350
If ($ExtOption.SiteList.Count/2 -is [int]){
    [int]$FormHeight = ((($ExtOption.SiteList.Count/2)*20)+255)   #--[ Dynamically Created Variable for Box Size (Even count) ]--
}Else{
    [int]$FormHeight = ((($ExtOption.SiteList.Count/2)*23)+255)   #--[ Dynamically Created Variable for Box Size (Odd count) ]--
}

[int]$FormHCenter = ($FormWidth / 2)   # 170 Horizontal center point
[int]$FormVCenter = ($FormHeight / 2)  # 209 Vertical center point
[int]$ButtonHeight = 25
[int]$TextHeight = 20

$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "FormWidth" -Value $FormWidth
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "FormHeight" -Value $FormHeight
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "FormHCenter" -Value $FormHCenter
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "FormVCenter" -Value $FormVCenter
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ButtonHeight" -Value $ButtonHeight
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TextHeight" -Value $TextHeight

StatusMsg ("Form Height = $FormHeight") "Magenta" $ExtOption
StatusMsg ("Form Width  = $FormWidth") "Magenta"  $ExtOption
StatusMsg ('Site Count  = '+$ExtOption.SiteList.Count) "Magenta" $ExtOption
StatusMsg ('Box Count   = '+($ExtOption.SiteList.Count/2)) "Magenta"  $ExtOption

#--[ Execute the form ]------------------------------------------------
If ($ExtOption.Console){
    $ExtOption
}

#$ExtOption    #--[ Uncomment to display the option object's contents ]--

ActivateForm $ExtOption

<#--[ Manual Shortcut Details ]---------------------------------------- 
To prevent any pop-up commend windows use the following in the "Target" field of a shortcut 
 
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -file "c:\scripts\Whereabouts.ps1" -windowstyle hidden -nonewwindow 
 
- Adjust the path to the script as needed. 
- Set the "Run" option to "Minimized" 
- An icon will appear briefly in the taskbar while assemblies load, then disappear as the GUI loads. 
 
#--[ Sample XML File ]--------------------------------------------------
<?xml version="1.0" encoding="utf-8"?>
<Settings>
    <General>
        <TitleText>Info Systems "Whereabouts" Manager</TitleText>	<!-- Title shown in script form header. -->
		<Shortcut>False</Shortcut>	                              <!-- Change to "True" to create a script shortcut automatically. -->
		<TravelLogPath> </TravelLogPath>	                      <!-- Defaults to users documents folder. -->
		<TravelLogFile>TravelLog.csv</TravelLogFile>	          <!-- Name of file to log miles to. -->
        <ZeroLog>False</ZeroLog>	                              <!-- Change to "True" to log zero mile destinations. -->
    </General>
	<Email>
        <SmtpServer>mailserver.my.org</SmtpServer>
        <SmtpPort>25</SmtpPort>
		<SSL>False</SSL>
		<HTML>True</HTML>
		<EmailUser></EmailUser>		                              <!-- Used if email authentication is needed. -->
		<EmailPass></EmailPass>	                                  <!-- Used if email authentication is needed. -->
		<!-- <SenderEmail>  This is auto-generated by the script   </SenderEmail> 
        <GroupEmail>InformationTechnology@my.org</GroupEmail>     <!-- This is the department/group email to send to -->
		<GroupEmailText>Infotech Dept</GroupEmailText>            <!-- This is displayed in-form by default for the group email -->
        <Recipients>                                              <!-- If no department/group email is entered this list will be used instead -->
            <Recipient>bob@my.org</Recipient>
            <Recipient>sam@my.org</Recipient>
            <Recipient>kathy@my.org</Recipient>
    <!--    <Recipient> Add more as needed </Recipient>    -->
        </Recipients>
    </Email>
    <Sites>
		<site>Hospital,5</site>		                              <!-- Format is "display name" a comma, then "mileage to site" -->
		<site>Datacenter,0</site> 
		<site>Site,1</site>
		<site>Clinic,2</site>
		<site>Office,6</site>
		<site>Other (see below),0</site>
    </Sites>
</Settings> 
















#>