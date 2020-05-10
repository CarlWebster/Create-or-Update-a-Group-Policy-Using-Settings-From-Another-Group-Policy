# Create-or-Update-a-Group-Policy-Using-Settings-From-Another-Group-Policy
Create an unlinked Group Policy Object from an imported CSV file

	Creates an unlinked Group Policy Object from an imported CSV file.
	
	If the Group Policy Object exists, linked or unlinked, it is updated with the 
	settings from the imported CSV file.
	
	The CSV file is created using Darren Mar-Elia's Registry .Pol Viewer Utility
	https://sdmsoftware.com/gpoguy/gpo-freeware-registration-form/

	The Registry.pol files only contain settings from the Administrative Templates
	node in the Group Policy Management Console.
	
	This script is designed for consultants and trainers who may create Group Policies 
	in a lab and need a way to recreate those policies at a customer or training site.
	
	The GroupPolicy module is required for the script to run.
	
	On a server, Group Policy Management must be installed.
	On a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
