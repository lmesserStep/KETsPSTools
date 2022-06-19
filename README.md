# KETsPSTools
Powershell scripts for Kentucky Education Technology - Active Directory Automation with GUI.
Please make sure that all fields in your CSV are populated. 


To begin, open 'Student Import Tool' on your local machine, do NOT run it from
A file server. It cannot load modules correctly when loaded from a server and will 
throw tons of errors. 

1.) Open Student Import Tool 
2.) Click 'Browse' and import your CSV. 
3.) Click connect and wait for connection, sometimes this takes 45 seconds. There is a time out for signing in. 
4.) Sign in when prompted.
5.) click 'Import' 

Wait 10-15 minutes before starting 'Email Enabler' 

1.) Open 'Email Enabler' 
2.) Click 'Import AD CMDs' 
3.) Browse for CSV 
4.) Check which boxes you want enabled. 'Email' will enable O365 and KETs Mail for the users. Description will populate description fields with passwords from CSV.

The program will NOT set email or descriptions for accounts that have a 1 in them. Example: john.doe1@stu.vendor.kyschools will not be found. This is a failsafe. These accounts will have to be manually enabled, but it will list them for you. 



*Do note, you will have to replace anything with **vendor** with your schools domain and domain controller.*
