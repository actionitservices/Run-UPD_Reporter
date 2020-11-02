The following script is will scan and report on User Profile Disks UVHDs via email.
Usernames and matching UVHDs are identified using the AD securityidentifier (SID).
User sessions must be logged in or UVHD mounted for volume health details.  
UPDPath will be scanned for *.vhdx files only.

EXAMPLE
 .\Run-UPD_Reporter.ps1 -To admin@example.com

NOTES 
File Name : .\Run-UPD_Reporter.ps1
Authors   : Neil Hennessy (Action IT Services)
Email:	   : info@actionitservices.com.au
Requires  : Windows Server and User Profile Disks (UVHD)

History
---------
Version   : 1.13 (13/12/17) - Initial Script 
Version   : 1.14 (16/12/17) - General Release