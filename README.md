# PWC Automation - PowerShell Script

## Script Functionality

### Values Extraction from PWC Excel file:

- Parses the given Excel file which was received from the PWC Team.
- Extract the values(IOCs) from all the sheets and make them into a list of values(IOCs).
- Skips the Legal Disclaimer sheet.

### Makes a new Excel file from extracted IOCs:

- Organizes the IOCs based on their types and also removes the Sanitization.
- Makes a new Excel file, organizes the IOCs in the created Excel file, and saves the new Excel file in the same location from where the script has been invoked.
- Names the new Excel file in the format of **PWC-dd-MM-yyyy HH:mm:ss**
  - Created file name will look like **[PWC-21-03-2-24 11:00:00.xlsx]**

### Note:

> To run this script the user must have installed `[Microsoft Office/Office]` with Excel on their machine.  
> As the Excel manipulation parts of this script completely rely on `Microsoft.Office.Interop.Excel` library.  
> The `FullPath` Argument will expect the complete path to the Excel file rather than the relative path.
> By default the Excel library will look and save the files in the "Documents" directory of the user machine.
> To avoid this, pass the full path to the Excel file while running the script.

# Instructions for running the script

### By default running scripts in any PowerShell is disabled as a security mechanism.

Read more about PowerShell Execution Policies [Here](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1)

> By default, the execution policy in Windows PowerShell is set to Restricted, which prevents users from running any scripts. This is a security measure to help protect users from malicious scripts.

- Advised way to set the PowerShell Execution policy is to RemoteSigned
- Below are the steps to set PowerShell Execution policy is to RemoteSigned.

## RemoteSigned

- The default execution policy for Windows server computers.
- Scripts can run.
- Requires a digital signature from a trusted publisher on scripts and configuration files that are downloaded from the internet which includes email and instant messaging programs.
- Doesn't require digital signatures on scripts that are written on the local computer and not downloaded from the internet.

> **NOTE: These Scripts are not Digitally Signed and the user must have the Admin permissions to change/set the Execution Policy to Remote Signed**

## _Step1:_

1. Change Windows execution policies to run scripts downloaded from the internet
2. Open PowerShell as Administrator
3. Check the current execution policy
   - Type `Get-ExecutionPolicy` and press Enter. This will usually display `"Restricted"` by default.
4. Change the execution policy
   - To allow running scripts downloaded from the internet, type `Set-ExecutionPolicy RemoteSigned` Press Enter, and confirm the change by typing "Y".
5. Run the below command  
   `.\pwc.ps1 -FullPath "Complete Path to the Excel Sheet"`

Example:
`.\pwc.ps1 -FullPath "D:\PATH\TO_YOUR\PWC_FILE.xlsx"`

## _Simpler Way:_

1. Run the command `powershell.exe -ExecutionPolicy ByPass -File .\pwc.ps1 -FullPath "Complete Path to the Excel Sheet"`
2. **ExecutionPolicy Bypass:** parameter tells PowerShell to temporarily bypass its default execution policy for this specific command.
3. By using "Bypass," you're instructing PowerShell to ignore any restrictions and run the script, even if it wouldn't normally be allowed.
4. The command essentially says, "Run the script named main.ps1, and while you're at it, ignore any execution policy restrictions that might normally prevent it from running."

Example:
`powershell.exe -ExecutionPolicy ByPass -File .\pwc.ps1 -FullPath "D:\PATH\TO_YOUR\PWC_FILE.xlsx"`
