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
                   : v7.10 - 05-16-25 - Corrected bug in siganture injection (HTML option missing in email function).  Fixed
                   :                    bug with infobox green color not applying during run.  Added neuter option for 
                   :                    quick testing.
                   : #>
                   $ScriptVer = "7.10"    <#--[ Current version # used in script ]--
                   : 
------------------------------------------------------------------------------#>
