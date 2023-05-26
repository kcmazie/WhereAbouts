Programatically creates an email to send to the predetermined user or group for notification about where you are.
Grabs Outlook email signature from current users profile.  Determines sender by current logged on user.  Optionally
writes a travel log to users Documents folder and optionally will lock the PC.  Dynamically grows or shrinks the
form depending on the number of sites in the site table.

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
                   : v6.00 - 05-05-23 - Relocated options out to XML file for publishing.  Refactored some sections.
