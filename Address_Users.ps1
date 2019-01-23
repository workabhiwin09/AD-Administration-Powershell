###########################################################
# AUTHOR  : Abhishek Maitra 
# DATE    : 20-06-2017 
# EDIT    : 20-06-2017
# COMMENT : This script Terminates Active Directory users,
#           including different kind of properties, based
#           on an input terminate_ad_users.csv.
# VERSION : 1.0
###########################################################
# CHANGELOG
# Version 1.2: 21-06-2017 
# - Changed the code for better
# - Added better Error Handling and Reporting.
# - Changed input file with more logical headers.
# - Added functionality for account Disabled,Description
# - Added the option to move every user to terminated OU.
# Version 1.2: 21-06-2017
###########################################################
Set-ExecutionPolicy RemoteSigned
#----------------------------------------------------------
#ERROR REPORTING ALL
#----------------------------------------------------------
Set-StrictMode -Version latest
#----------------------------------------------------------
# LOAD ASSEMBLIES AND MODULES
#----------------------------------------------------------

Try
{
  Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
  Write-Host "[ERROR]`t ActiveDirectory Module couldn't be loaded. Script will stop!"
  Exit 1
}

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
#Give the Stored CSV file path
$newpath  = "path to terminate ad users csv file"
#Give the path where logs will be stored
$log      = "path to add ad user terminated logs"
$date     = Get-Date
$i        = 0
$Disable_Date = Read-Host -Prompt "Enter the Disable date in MMYY format"

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------
Function Start-Commands
{
  Terminate-Users
}

Function Terminate-Users
{
  "Processing started (on " + $date + "): " | Out-File $log -append
  "--------------------------------------------" | Out-File $log -append
  Import-CSV $newpath | ForEach-Object {
    If (($_.Term.ToLower()) -eq "yes")
    {
      If (($_.Global_ID -eq "") -Or ($_.CA_RequestNo -eq "") )
      {
        Write-Host "[ERROR]`t Please provide valid Global_ID and RequestNo. Processing skipped for line $($i)`r`n"
        "[ERROR]`t Please provide valid Global_ID and RequestNo. Processing skipped for line $($i)`r`n" | Out-File $log -append
      }
      Else
      {
                        
        # Set Description and Initials
        $Desc =$Disable_Date +"-ABH " + $_.Description + $_.RequestNo
        
        #Store the Global Id and the email Address
        $sam = $_.Global_ID
        $email = $_.EmailAddr

        Try
          {
            Write-Host "[INFO]`t Terminating User : $($sam) -- $($email)"
            "[INFO]`t Terminating User : $($sam) -- $($email)" | Out-File $log -append
            
            #Disable AD Account
            Disable-ADAccount $sam
        
            #Set AD Description
            Set-ADUser $sam -Description $Desc

            #Move to terminate OU
            # Set the target OU
            $location = (Get-ADOrganizationalUnit -LDAPFilter "(name=Terminated Users)").DistinguishedName
            $dn = (Get-ADUser $sam).DistinguishedName
            Move-ADObject -Identity $dn -TargetPath $location
            
            Write-Host "[INFO]`t Terminated User Successfully : $($sam)`t"
            "[INFO]`t Terminated User Successfully : $($sam)`t" | Out-File $log -append
          }
 
         Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
          }
        }
      }
     Else
    {
      Write-Host "[SKIP]`t User ($($_.Global_ID) $($_.EmailAddr)) will be skipped for processing!`r`n"
      "[SKIP]`t User ($($_.Global_ID) $($_.EmailAddr)) will be skipped for processing!" | Out-File $log -append
    }
    $i++
  }
  "--------------------------------------------" + "`r`n" | Out-File $log -append
  Write-Host "Total Users terminated are " $i
 }
Write-Host "STARTED SCRIPT`r`n"
Start-Commands
Write-Host "STOPPED SCRIPT"