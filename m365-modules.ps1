[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(785,450)
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow #modifies the window border
$Form.Text = "Microsoft PoweShell Installation Toool - www.bonguides.com" #window description
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

$precheck = {
   if ((Get-ExecutionPolicy) -notmatch "RemoteSigned") {
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
}
if (((Get-PSRepository -Name PSGallery).InstallationPolicy) -match "Untrusted") {
   Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
}
}

$installmsoline = {
   Write-Host
   Write-Host "Installing MSOnline Module..."
   Write-Host
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name MSOnline
   Write-Host
   Write-Host "Done."
   Write-Host
}
$installazuread = { 
   Write-Host
   Write-Host "Installing Azure AD Module..."
   Write-Host
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module AzureAD -Force
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installexchangeonline = { 
   Write-Host
   Write-Host "Installing Exchange Online Module..."
   Write-Host
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name ExchangeOnlineManagement -Force
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installmsteams = {
   Write-Host
   Write-Host "Installing Microsoft Teams Module..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name MicrosoftTeams -Force -AllowClobber
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installsharepoint = { 
   Write-Host
   Write-Host "Installing SharePoint Online Module..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installgraphapi = {
   Write-Host
   Write-Host "Installing Microsoft Graph Module..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module Microsoft.Graph -Force
   Install-Module MSAL.PS -Scope AllUsers -Force   
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installazureaz = {
   Write-Host
   Write-Host "Installing Azure Az Modules..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name Az -Force  
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installmde = {
   Write-Host
   Write-Host "Installing Microsoft Defender for Endpoint Modules..."
   Write-Host  
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name "PSMDATP"
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installintune = {
   Write-Host
   Write-Host "Installing Microsoft Graph Intune Module..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name Microsoft.Graph.Intune
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installuprint = { 
   Write-Host
   Write-Host "Installing Universal Print Management Module..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module UniversalPrintManagement
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installpowerapps = { 
   Write-Host
   Write-Host "Installing Microsoft PowerApps Modules..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force
   Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Force
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installazurerm = {
   Write-Host
   Write-Host "Installing AzureRM Modules..."
   Write-Host 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Install-Module -Name AzureRM -Scope CurrentUser -Repository PSGallery -Force
   Write-Host
   Write-Host "Done."
   Write-Host
}

$installall = { 
   Install-PackageProvider -Name NuGet -Force
   Invoke-Command $precheck
   Write-Host
   Write-Host "(1/12) Installing MSOnline Module..."
   Write-Host
   Install-Module -Name MSOnline
   Write-Host
   Write-Host "(2/12) Installing AzureAD Module..."
   Write-Host
   Install-Module AzureAD -Force
   Write-Host
   Write-Host "(3/12) Installing Exchange Online Management Module..."
   Write-Host
   Install-Module -Name ExchangeOnlineManagement -Force
   Write-Host
   Write-Host "(4/12) Installing Microsoft Teams Module..."
   Write-Host
   Install-Module -Name MicrosoftTeams -Force -AllowClobber
   Write-Host
   Write-Host "(5/12) Installing SharePoint Online Module..."
   Write-Host
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
   Write-Host
   Write-Host "(6/12) Installing Microsoft Graph Modules..."
   Write-Host
   Install-Module Microsoft.Graph -Force
   Install-Module MSAL.PS -Force 
   Write-Host
   Write-Host "(7/12) Installing Azure Az Module..."
   Write-Host
   Install-Module -Name Az -Force 
   Write-Host
   Write-Host "(8/12) Installing Microsoft Defender for Endpoint Module..."
   Write-Host
   Install-Module -Name "PSMDATP"
   Write-Host
   Write-Host "(9/12) Installing Microsoft Graph Intune Module..."
   Write-Host
   Install-Module -Name Microsoft.Graph.Intune
   Write-Host
   Write-Host "(10/12) Installing Universal Print Management Module..."
   Write-Host
   Install-Module -Name UniversalPrintManagement
   Write-Host
   Write-Host "(11/12) Installing Microsoft PowerApps Modules..."
   Write-Host
   Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force
   Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Force
   Write-Host
   Write-Host "(12/12) Installing AzureRM Modules..."
   Write-Host
   Install-Module -Name AzureRM -Scope CurrentUser -Repository PSGallery -Force
   Write-Host
   Write-Host "Done, all modules have been installed."
   Write-Host
}


############################################## Start functions

   function m365PowerShell {
   try {
   if ($msonlinecb.Checked -eq $true) {Invoke-Command $installmsoline}
   if ($azureadcb.Checked -eq $true) {Invoke-Command $installazuread}
   if ($exchangeonlinecb.Checked -eq $true) {Invoke-Command $installexchangeonline}
   if ($sharepointcb.Checked -eq $true) {Invoke-Command $installsharepoint}
   if ($graphapicb.Checked -eq $true) {Invoke-Command $installgraphapi}
   if ($azureazcb.Checked -eq $true) {Invoke-Command $installazureaz}
   if ($mdecb.Checked -eq $true) {Invoke-Command $installmde}
   if ($intunecb.Checked -eq $true) {Invoke-Command $installintune}
   if ($uprintcb.Checked -eq $true) {Invoke-Command $installuprint}
   if ($azurermcb.Checked -eq $true) {Invoke-Command $installazurerm}
   if ($powerappscb.Checked -eq $true) {Invoke-Command $installpowerapps}
   if ($allcb.Checked -eq $true) {Invoke-Command $installall}

   } #end try

   catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

############################################## Start group boxes

   $installModule = New-Object System.Windows.Forms.GroupBox
   $installModule.Location = New-Object System.Drawing.Size(10,10) 
   $installModule.size = New-Object System.Drawing.Size(130,270) 
   $installModule.text = "Install the Module:"
   $Form.Controls.Add($installModule) 

   $connectTo = New-Object System.Windows.Forms.GroupBox
   $connectTo.Location = New-Object System.Drawing.Size(150,10) 
   $connectTo.size = New-Object System.Drawing.Size(130,220) 
   $connectTo.text = "Connect:"
   $Form.Controls.Add($connectTo) 

############################################## end group boxes

############################################## Start Arch checkboxes

   $msonlinecb = New-Object System.Windows.Forms.RadioButton
   $msonlinecb.Location = New-Object System.Drawing.Size(10,20)
   $msonlinecb.Size = New-Object System.Drawing.Size(100,20)
   $msonlinecb.Checked = $false
   $msonlinecb.Text = "MSOnline"
   $installModule.Controls.Add($msonlinecb)

   $azureadcb = New-Object System.Windows.Forms.RadioButton
   $azureadcb.Location = New-Object System.Drawing.Size(10,40)
   $azureadcb.Size = New-Object System.Drawing.Size(100,20)
   $azureadcb.Checked = $false
   $azureadcb.Text = "AzureAD"
   $installModule.Controls.Add($azureadcb)

   $exchangeonlinecb = New-Object System.Windows.Forms.RadioButton
   $exchangeonlinecb.Location = New-Object System.Drawing.Size(10,60)
   $exchangeonlinecb.Size = New-Object System.Drawing.Size(100,20)
   $exchangeonlinecb.AutoSize = $True
   $exchangeonlinecb.Text = "Exchange Online"
   $installModule.Controls.Add($exchangeonlinecb)

   $msteamscb = New-Object System.Windows.Forms.RadioButton
   $msteamscb.Location = New-Object System.Drawing.Size(10,80)
   $msteamscb.Size = New-Object System.Drawing.Size(100,20)
   $msteamscb.Checked = $false
   $msteamscb.Text = "Teams"
   $installModule.Controls.Add($msteamscb)

   $sharepointcb = New-Object System.Windows.Forms.RadioButton
   $sharepointcb.Location = New-Object System.Drawing.Size(10,100)
   $sharepointcb.Size = New-Object System.Drawing.Size(100,20)
   $sharepointcb.Checked = $false
   $sharepointcb.Text = "Sharepoint"
   $installModule.Controls.Add($sharepointcb)

   $azureazcb = New-Object System.Windows.Forms.RadioButton
   $azureazcb.Location = New-Object System.Drawing.Size(10,120)
   $azureazcb.Size = New-Object System.Drawing.Size(100,20)
   $azureazcb.Text = "Azure Az"
   $installModule.Controls.Add($azureazcb)

   $azurermcb = New-Object System.Windows.Forms.RadioButton
   $azurermcb.Location = New-Object System.Drawing.Size(10,120)
   $azurermcb.Size = New-Object System.Drawing.Size(100,20)
   $azurermcb.Text = "Azure RM"
   $installModule.Controls.Add($azurermcb)

   $mdecb = New-Object System.Windows.Forms.RadioButton
   $mdecb.Location = New-Object System.Drawing.Size(10,140)
   $mdecb.Size = New-Object System.Drawing.Size(100,20)
   $mdecb.AutoSize = $True
   $mdecb.Text = "Defender Endpoint"
   $installModule.Controls.Add($mdecb)

   $intunecb = New-Object System.Windows.Forms.RadioButton
   $intunecb.Location = New-Object System.Drawing.Size(10,160)
   $intunecb.Size = New-Object System.Drawing.Size(100,20)
   $intunecb.Text = "Intune"
   $installModule.Controls.Add($intunecb)

   $uprintcb = New-Object System.Windows.Forms.RadioButton
   $uprintcb.Location = New-Object System.Drawing.Size(10,180)
   $uprintcb.Size = New-Object System.Drawing.Size(100,20)
   $uprintcb.Text = "Universal Print"
   $installModule.Controls.Add($uprintcb)

   $powerappscb = New-Object System.Windows.Forms.RadioButton
   $powerappscb.Location = New-Object System.Drawing.Size(10,200)
   $powerappscb.Size = New-Object System.Drawing.Size(100,20)
   $powerappscb.Text = "PowerApps"
   $installModule.Controls.Add($powerappscb)

   $graphapicb = New-Object System.Windows.Forms.RadioButton
   $graphapicb.Location = New-Object System.Drawing.Size(10,220)
   $graphapicb.Size = New-Object System.Drawing.Size(100,20)
   $graphapicb.Text = "Grap API"
   $installModule.Controls.Add($graphapicb)

   $allcb = New-Object System.Windows.Forms.RadioButton
   $allcb.Location = New-Object System.Drawing.Size(10,240)
   $allcb.Size = New-Object System.Drawing.Size(100,20)
   $allcb.Text = "All Services"
   $installModule.Controls.Add($allcb)
############################################## End Arch checkboxes 

############################################## Start buttons

   $submitButton = New-Object System.Windows.Forms.Button 
   $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitButton.Location = New-Object System.Drawing.Size(10,280) 
   $submitButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitButton.Text = "Submit" 
   $submitButton.Add_Click({m365PowerShell}) 
   $Form.Controls.Add($submitButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
