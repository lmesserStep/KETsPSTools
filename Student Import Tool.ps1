<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
     Exchange Tool by Lmesser
#>

param([switch]$Elevated)

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) 
    {
        # tried to elevate, did not work, aborting
    } 
    else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -ExecutionPolicy Bypass -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
}

exit
}

'running with full privileges'

$ErrorActionPreference = 'SilentlyContinue'

$Button = [System.Windows.MessageBoxButton]::YesNoCancel
$ErrorIco = [System.Windows.MessageBoxImage]::Error
$Ask = 'Do you want to run this as an Administrator?
        Select "Yes" to Run as an Administrator
        Select "No" to not run this as an Administrator
        
        Select "Cancel" to stop the script.'

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    $Prompt = [System.Windows.MessageBox]::Show($Ask, "Run as an Administrator or not?", $Button, $ErrorIco) 
    Switch ($Prompt) {
        #This will debloat Windows 10
        Yes {
            Write-Host "You didn't run this script as an Administrator. This script will self elevate to run as an Administrator and continue."
            Start-Process PowerShell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
            Exit

        }
        No {

            Break
        }
    }
}


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(646,277)
$Form.text                       = "Exchange Student Import Tool By Lmesserfor KETS Schools"
$Form.TopMost                    = $false

$Connect                         = New-Object system.Windows.Forms.Button
$Connect.text                    = "Connect"
$Connect.width                   = 82
$Connect.height                  = 35
$Connect.location                = New-Object System.Drawing.Point(13,9)
$Connect.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 178
$ProgressBar1.height             = 23
$ProgressBar1.location           = New-Object System.Drawing.Point(12,240)

$Disconnect                      = New-Object system.Windows.Forms.Button
$Disconnect.text                 = "Disconnect"
$Disconnect.width                = 84
$Disconnect.height               = 38
$Disconnect.enabled              = $false
$Disconnect.location             = New-Object System.Drawing.Point(13,52)
$Disconnect.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Information                     = New-Object system.Windows.Forms.Label
$Information.text                = "Click `'Connect`' to begin"
$Information.AutoSize            = $true
$Information.width               = 25
$Information.height              = 10
$Information.location            = New-Object System.Drawing.Point(13,208)
$Information.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Browse                          = New-Object system.Windows.Forms.Button
$Browse.text                     = "Browse"
$Browse.width                    = 60
$Browse.height                   = 30
$Browse.location                 = New-Object System.Drawing.Point(164,9)
$Browse.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 217
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(229,16)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Import                          = New-Object system.Windows.Forms.Button
$Import.text                     = "Start"
$Import.width                    = 82
$Import.height                   = 47
$Import.enabled                  = $false
$Import.location                 = New-Object System.Drawing.Point(549,210)
$Import.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "CSV Example"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(271,66)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "firstname"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(61,96)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "lastname"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(129,96)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "name"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(197,96)
$Label4.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "smtp"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(253,96)
$Label5.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "userprinc"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(296,96)
$Label6.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "password"
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(368,96)
$Label7.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label8                          = New-Object system.Windows.Forms.Label
$Label8.text                     = "remoteaddress"
$Label8.AutoSize                 = $true
$Label8.width                    = 25
$Label8.height                   = 10
$Label8.location                 = New-Object System.Drawing.Point(434,96)
$Label8.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label9                          = New-Object system.Windows.Forms.Label
$Label9.text                     = "desc"
$Label9.AutoSize                 = $true
$Label9.width                    = 25
$Label9.height                   = 10
$Label9.location                 = New-Object System.Drawing.Point(536,96)
$Label9.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label10                         = New-Object system.Windows.Forms.Label
$Label10.text                    = "user"
$Label10.AutoSize                = $true
$Label10.width                   = 25
$Label10.height                  = 10
$Label10.location                = New-Object System.Drawing.Point(584,96)
$Label10.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($Connect,$ProgressBar1,$Disconnect,$Information,$Browse,$TextBox1,$Import,$Label1,$Label2,$Label3,$Label4,$Label5,$Label6,$Label7,$Label8,$Label9,$Label10))




#Write your logic code here
#Made by lmesser Messer District Technology Cordinator vendor County School System


$Connect.Add_Click( {

      
       winrm set winrm/config/client/auth '@{Basic="true"}'
       
       $progressbar1.Value =20; 
       $Information.text = "Setting Winrm config auth to basic...";
       Start-Sleep -Seconds 5
       
       Import-Module ActiveDirectory 
       
       $progressbar1.Value =37; 
       $Information.text = "Importing Active Directory Modules...";
       
       Start-Sleep -Seconds 10 
       
       $progressbar1.Value =78; 
       $Information.text = "Connecting to vendor.ketsds.net through mail.education.ky.gov...";
       $UserCredential = Get-Credential
       
       Start-Sleep -Seconds 4
       $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://mail.education.ky.gov/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection
       
       Start-Sleep -Seconds 15 
       
       $progressbar1.Value =89; 
       $Information.text = "Connected! Importing session to local machine...";

       Import-PSSession $Session -DisableNameChecking
       
       
      Start-Sleep -Seconds 15 
      
      $progressbar1.Value =98; 
      $Information.text = "Session imported successfully... You are now connected to vendor.kets.net and ED031ADGC1.vendor.ketsds.net";
       
       Start-Sleep -Seconds 5 
      
       $progressbar1.Value =0; 
      $Information.text = "Connected.";
	  
If ( $Information.text -eq "Connected."){
        $Connect.enabled = $false 
        $Connect.text = "Connected"
        $Disconnect.enabled = $true
        $Import.enabled = $true
        Write-Host "IT IS VERY IMPORTANT TO DISCONNECT BEFORE CLOSING!!!!"
      }else{
          $Disconnect.enabled = $false
		   }
 
}
      
)




$Browse.Add_Click( {
      add-type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.StartPosition = 'CenterScreen'

$button = New-Object System.Windows.Forms.Button
$form.Controls.Add($button)
$button.Text = 'Get file'
$button.Location = '10,10'
$button.add_Click({
    $ofd = New-Object system.windows.forms.Openfiledialog
    $ofd.Filter =  'CSVs (*.csv)|*.csv' 
    $script:filename = 'Not found'
    if($ofd.ShowDialog() -eq 'Ok'){
        $script:filename = $textbox.Text = $ofd.FileName
    }    
})

$buttonOK = New-Object System.Windows.Forms.Button
$form.Controls.Add($buttonOK)
$buttonOK.Text = 'Okay'
$buttonOK.Location = '10,40'
$buttonOK.DialogResult = 'OK'
$buttonOK.Add_Click( {
    $Information.Text = "File loaded" 
    }
    )

$textbox = New-Object System.Windows.Forms.TextBox
$form.Controls.Add($textbox)
$textbox.Location = '100,10'
$textbox.Width += 50

$form.ShowDialog()
$filename
$textbox1.text = $filename 
    }
)

$Import.Add_Click( {
       
            Import-CSV -path $filename | ForEach {New-RemoteMailbox -FirstName $_.firstname -LastName $_.lastname -Name $_.name -PrimarySmtpAddress $_.smtp -UserPrincipalName $_.userprinc -Password (ConvertTo-SecureString $_.password -AsPlainText -Force) -OnPremisesOrganizationalUnit "vendor.ketsds.net/Students" -RemoteroutingAddress $_.remoteaddress -DomainController ED031ADGC1.vendor.ketsds.net} 
            $Information.Text = "$filename accounts imported successfully."
            $Group.enabled = $true
            $Groups.enabled = $true 
            $Import.enabled = $false 
            Start-Sleep -Seconds 25 
            $Information.text = "Please wait for accounts to populate..."
            $Information.text = "Disconnect and use 'Populate Mail and Description' tool."
            
            
        }
)



$Disconnect.Add_Click( {
   Remove-PSSession 1 
   $Information.text = "Session has been disconnected successfully."
   $Connect.Text ="Connect"
   $Connect.enabled = "True"
   $Disconnect.enabled = "False"
   
   
    }
)



[void]$Form.ShowDialog()
