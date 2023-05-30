[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(900,400)
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow #modifies the window border
$Form.Text = "Microsoft Office Installation Toool - www.bonguides.com" #window description
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

############################################## Start functions

function WinGetInstaller {
try {
   
$winGet = gci "C:\Program Files\WindowsApps" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.name -like "AppInstallerCLI.exe" -or $_.name -like "WinGet.exe" } | Select-Object -ExpandProperty fullname -ErrorAction SilentlyContinue

# If there are multiple versions, select latest
if ($winGet.count -gt 1){
    $winGet = $winGet[-1]
}
$winGetLoc = [string]((get-item $winGet).Directory.FullName)

$install = {
    Write-Host "Installing $id ..."
    & "$winGetLoc\winget.exe" install --id=$id -e --silent --accept-source-agreements --accept-package-agreements
}   
    
$ids=@()
    if ($chromecb.Checked -eq $true) {$ids += 'Google.Chrome'}
    if ($firefoxcb.Checked -eq $true) {$ids += 'Mozilla.Firefox'}
    if ($bravecb.Checked -eq $true) {$ids += 'Brave.Brave'}
    if ($edgecb.Checked -eq $true) {$ids += 'Microsoft.Edge'}

    if ($discordcb.Checked -eq $true) {$ids += 'Discord.Discord'}
    if ($skypecb.Checked -eq $true) {$ids += 'Microsoft.Skype'}
    if ($slackcb.Checked -eq $true) {$ids += 'SlackTechnologies.Slack'}
    if ($telegramcb.Checked -eq $true) {$ids += 'Telegram.TelegramDesktop'}
    if ($teamscb.Checked -eq $true) {$ids += 'Microsoft.Teams'}
    if ($vibercb.Checked -eq $true) {$ids += 'Viber.Viber'}
    if ($zoomcb.Checked -eq $true) {$ids += 'Zoom.Zoom'}

    if ($puttycb.Checked -eq $true) {$ids += 'PuTTY.PuTTY'}    
    if ($winscpcb.Checked -eq $true) {$ids += 'WinSCP.WinSCP'}    
    if ($tcpviewcb.Checked -eq $true) {$ids += 'Microsoft.Sysinternals.TCPView'}    
    if ($mremotengcb.Checked -eq $true) {$ids += 'mRemoteNG.mRemoteNG'}    
    if ($ipscancb.Checked -eq $true) {$ids += 'Famatech.AdvancedIPScanner'}    
    if ($wiresharkcb.Checked -eq $true) {$ids += 'WiresharkFoundation.Wireshark'}    

    if ($vlcplayercb.Checked -eq $true) {$ids += 'VideoLAN.VLC'}    
    if ($intunecb.Checked -eq $true) {$ids += 'Apple.iTunes'}    
    if ($obsstudiocb.Checked -eq $true) {$ids += 'OBSProject.OBSStudio'}    
    if ($sharexcb.Checked -eq $true) {$ids += 'ShareX.ShareX'}    
    if ($handbreakcb.Checked -eq $true) {$ids += 'HandBrake.HandBrake'}    
    if ($flameshotcb.Checked -eq $true) {$ids += 'Flameshot.Flameshot'}    





ForEach ($id in $ids){
Invoke-Command $install
}

Write-Host "Done."

} #end try

catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

function WinGetUninstaller {

    try {

    $winGet = gci "C:\Program Files\WindowsApps" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.name -like "AppInstallerCLI.exe" -or $_.name -like "WinGet.exe" } | Select-Object -ExpandProperty fullname -ErrorAction SilentlyContinue

    # If there are multiple versions, select latest
    if ($winGet.count -gt 1){
        $winGet = $winGet[-1]
    }
    $winGetLoc = [string]((get-item $winGet).Directory.FullName)

    $uninstall = {
    Write-Host "Removing $uid ..."
    & "$winGetLoc\winget.exe" uninstall --id=$uid -e --silent 
    }   
        
    $uids=@()
    if ($chromecb.Checked -eq $true) {$uids += 'Google.Chrome'}
    if ($firefoxcb.Checked -eq $true) {$uids += 'Mozilla.Firefox'}
    if ($bravecb.Checked -eq $true) {$uids += 'Brave.Brave'}
    if ($edgecb.Checked -eq $true) {$uids += 'Microsoft.Edge'}

    ForEach ($uid in $uids){
    Invoke-Command $uninstall
    }

    Write-Host "Done."

    } #end try

    catch {$outputBox.text = "`nOperation could not be completed"}

    } 
############################################## end functions

############################################## Start group boxes

   $Browsers = New-Object System.Windows.Forms.GroupBox
   $Browsers.Location = New-Object System.Drawing.Size(10,10) 
   $Browsers.size = New-Object System.Drawing.Size(130,250) 
   $Browsers.text = "Browsers:"
   $Form.Controls.Add($Browsers) 

   $Comunications = New-Object System.Windows.Forms.GroupBox
   $Comunications.Location = New-Object System.Drawing.Size(150,10) 
   $Comunications.size = New-Object System.Drawing.Size(130,250) 
   $Comunications.text = "Comunications:"
   $Form.Controls.Add($Comunications) 

   $NetworkTools = New-Object System.Windows.Forms.GroupBox
   $NetworkTools.Location = New-Object System.Drawing.Size(290,10) 
   $NetworkTools.size = New-Object System.Drawing.Size(130,250) 
   $NetworkTools.text = "Network Tools:"
   $Form.Controls.Add($NetworkTools) 

   $Multimedia = New-Object System.Windows.Forms.GroupBox
   $Multimedia.Location = New-Object System.Drawing.Size(430,10) 
   $Multimedia.size = New-Object System.Drawing.Size(130,250) 
   $Multimedia.text = "Multimedia:"
   $Form.Controls.Add($Multimedia) 

   $Document = New-Object System.Windows.Forms.GroupBox
   $Document.Location = New-Object System.Drawing.Size(570,10) 
   $Document.size = New-Object System.Drawing.Size(130,250) 
   $Document.text = "Document:"
   $Form.Controls.Add($Document)

   $Games = New-Object System.Windows.Forms.GroupBox
   $Games.Location = New-Object System.Drawing.Size(710,10) 
   $Games.size = New-Object System.Drawing.Size(130,250) 
   $Games.text = "Games:"
   $Form.Controls.Add($Games)

   $Utilities = New-Object System.Windows.Forms.GroupBox
   $Utilities.Location = New-Object System.Drawing.Size(850,10) 
   $Utilities.size = New-Object System.Drawing.Size(130,250) 
   $Utilities.text = "Utilities:"
   $Form.Controls.Add($Utilities)

############################################## end group boxes

############################################## Start Browsers checkboxes

   $chromecb = New-Object System.Windows.Forms.checkbox
   $chromecb.Location = New-Object System.Drawing.Size(10,20)
   $chromecb.Size = New-Object System.Drawing.Size(100,20)
   $chromecb.Text = "Chrome"
   $chromecb.AutoSize = $True
   $Browsers.Controls.Add($chromecb)

   $firefoxcb = New-Object System.Windows.Forms.checkbox
   $firefoxcb.Location = New-Object System.Drawing.Size(10,40)
   $firefoxcb.Size = New-Object System.Drawing.Size(100,20)
   $firefoxcb.Text = "Firefox"
   $Browsers.Controls.Add($firefoxcb)

   $bravecb = New-Object System.Windows.Forms.checkbox
   $bravecb.Location = New-Object System.Drawing.Size(10,60)
   $bravecb.Size = New-Object System.Drawing.Size(100,20)
   $bravecb.Text = "Brave"
   $Browsers.Controls.Add($bravecb)

   $egdecb = New-Object System.Windows.Forms.checkbox
   $egdecb.Location = New-Object System.Drawing.Size(10,80)
   $egdecb.Size = New-Object System.Drawing.Size(100,20)
   $egdecb.Text = "Edge"
   $Browsers.Controls.Add($egdecb)

   $torcb = New-Object System.Windows.Forms.checkbox
   $torcb.Location = New-Object System.Drawing.Size(10,100)
   $torcb.Size = New-Object System.Drawing.Size(100,20)
   $torcb.Text = "Tor Browser"
   $Browsers.Controls.Add($torcb)

   $vivaldicb = New-Object System.Windows.Forms.checkbox
   $vivaldicb.Location = New-Object System.Drawing.Size(10,120)
   $vivaldicb.Size = New-Object System.Drawing.Size(100,20)
   $vivaldicb.Text = "Vivaldi"
   $Browsers.Controls.Add($vivaldicb)

############################################## End Arch checkboxes

############################################## Start Communications checkboxes 

   $discordcb = New-Object System.Windows.Forms.checkbox
   $discordcb.Location = New-Object System.Drawing.Size(10,20)
   $discordcb.Size = New-Object System.Drawing.Size(100,20)
   $discordcb.Text = "Discord"
   $discordcb.AutoSize = $True
   $Comunications.Controls.Add($discordcb)

   $skypecb = New-Object System.Windows.Forms.checkbox
   $skypecb.Location = New-Object System.Drawing.Size(10,40)
   $skypecb.Size = New-Object System.Drawing.Size(100,20)
   $skypecb.Text = "Skype"
   $skypecb.AutoSize = $True
   $Comunications.Controls.Add($skypecb)

   $slackcb = New-Object System.Windows.Forms.checkbox
   $slackcb.Location = New-Object System.Drawing.Size(10,60)
   $slackcb.Size = New-Object System.Drawing.Size(100,20)
   $slackcb.Text = "Slack"
   $slackcb.AutoSize = $True
   $Comunications.Controls.Add($slackcb)

   $telegramcb = New-Object System.Windows.Forms.checkbox
   $telegramcb.Location = New-Object System.Drawing.Size(10,80)
   $telegramcb.Size = New-Object System.Drawing.Size(100,20)
   $telegramcb.Text = "Telegram"
   $telegramcb.AutoSize = $True
   $Comunications.Controls.Add($telegramcb)

   $teamscb = New-Object System.Windows.Forms.checkbox
   $teamscb.Location = New-Object System.Drawing.Size(10,100)
   $teamscb.Size = New-Object System.Drawing.Size(100,20)
   $teamscb.Text = "Teams"
   $teamscb.AutoSize = $True
   $Comunications.Controls.Add($teamscb)

   $vibercb = New-Object System.Windows.Forms.checkbox
   $vibercb.Location = New-Object System.Drawing.Size(10,120)
   $vibercb.Size = New-Object System.Drawing.Size(100,20)
   $vibercb.Text = "Viber"
   $vibercb.AutoSize = $True
   $Comunications.Controls.Add($vibercb)

   $zoomcb = New-Object System.Windows.Forms.checkbox
   $zoomcb.Location = New-Object System.Drawing.Size(10,140)
   $zoomcb.Size = New-Object System.Drawing.Size(100,20)
   $zoomcb.Text = "Zoom"
   $zoomcb.AutoSize = $True
   $Comunications.Controls.Add($zoomcb)

############################################## End Communications checkboxes 


############################################## Start Network Tools checkboxes 

   $puttycb = New-Object System.Windows.Forms.checkbox
   $puttycb.Location = New-Object System.Drawing.Size(10,20)
   $puttycb.Size = New-Object System.Drawing.Size(100,20)
   $puttycb.Text = "PuTTy"
   $puttycb.AutoSize = $True
   $NetworkTools.Controls.Add($puttycb)

   $winscpcb = New-Object System.Windows.Forms.checkbox
   $winscpcb.Location = New-Object System.Drawing.Size(10,40)
   $winscpcb.Size = New-Object System.Drawing.Size(100,20)
   $winscpcb.Text = "WinSCP"
   $NetworkTools.Controls.Add($winscpcb)

   $tcpviewcb = New-Object System.Windows.Forms.checkbox
   $tcpviewcb.Location = New-Object System.Drawing.Size(10,60)
   $tcpviewcb.Size = New-Object System.Drawing.Size(100,20)
   $tcpviewcb.Text = "TCP View"
   $NetworkTools.Controls.Add($tcpviewcb)

   $mremotengcb = New-Object System.Windows.Forms.checkbox
   $mremotengcb.Location = New-Object System.Drawing.Size(10,80)
   $mremotengcb.Size = New-Object System.Drawing.Size(100,20)
   $mremotengcb.Text = "mRemoteNG"
   $NetworkTools.Controls.Add($mremotengcb)

   $ipscancb = New-Object System.Windows.Forms.checkbox
   $ipscancb.Location = New-Object System.Drawing.Size(10,100)
   $ipscancb.Size = New-Object System.Drawing.Size(100,20)
   $ipscancb.Text = "IP Scanner"
   $NetworkTools.Controls.Add($ipscancb)

   $wiresharkcb = New-Object System.Windows.Forms.checkbox
   $wiresharkcb.Location = New-Object System.Drawing.Size(10,120)
   $wiresharkcb.Size = New-Object System.Drawing.Size(100,20)
   $wiresharkcb.Text = "WireShark"
   $NetworkTools.Controls.Add($wiresharkcb)            

############################################## End Network Tools  checkboxes 


############################################## Start Multimedia checkboxes 

   $vlcplayercb = New-Object System.Windows.Forms.checkbox
   $vlcplayercb.Location = New-Object System.Drawing.Size(10,20)
   $vlcplayercb.Size = New-Object System.Drawing.Size(100,20)
   $vlcplayercb.Text = "VLC Player"
   $vlcplayercb.AutoSize = $True
   $Multimedia.Controls.Add($vlcplayercb)

   $intunecb = New-Object System.Windows.Forms.checkbox
   $intunecb.Location = New-Object System.Drawing.Size(10,40)
   $intunecb.Size = New-Object System.Drawing.Size(100,20)
   $intunecb.Text = "iTunes"
   $Multimedia.Controls.Add($intunecb)

   $obsstudiocb = New-Object System.Windows.Forms.checkbox
   $obsstudiocb.Location = New-Object System.Drawing.Size(10,60)
   $obsstudiocb.Size = New-Object System.Drawing.Size(100,20)
   $obsstudiocb.Text = "OBS Studio"
   $Multimedia.Controls.Add($obsstudiocb)

   $sharexcb = New-Object System.Windows.Forms.checkbox
   $sharexcb.Location = New-Object System.Drawing.Size(10,80)
   $sharexcb.Size = New-Object System.Drawing.Size(100,20)
   $sharexcb.Text = "ShareX"
   $Multimedia.Controls.Add($sharexcb)

   $handbreakcb = New-Object System.Windows.Forms.checkbox
   $handbreakcb.Location = New-Object System.Drawing.Size(10,100)
   $handbreakcb.Size = New-Object System.Drawing.Size(100,20)
   $handbreakcb.Text = "HandBrake"
   $Multimedia.Controls.Add($handbreakcb)

   $flameshotcb = New-Object System.Windows.Forms.checkbox
   $flameshotcb.Location = New-Object System.Drawing.Size(10,120)
   $flameshotcb.Size = New-Object System.Drawing.Size(100,20)
   $flameshotcb.Text = "FlameShot"
   $Multimedia.Controls.Add($flameshotcb)



############################################## End Multimedia checkboxes 



############################################## Start Document checkboxes 
############################################## End Document checkboxes 


############################################## Start Games checkboxes 
############################################## End Games checkboxes 


############################################## Start Ultilities checkboxes 
############################################## End Ultilities checkboxes 


############################################## Start buttons

   $submitInstallButton = New-Object System.Windows.Forms.Button 
   $submitInstallButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitInstallButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitInstallButton.Location = New-Object System.Drawing.Size(10,280) 
   $submitInstallButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitInstallButton.Text = "Install " 
   $submitInstallButton.Add_Click({WinGetInstaller}) 
   $Form.Controls.Add($submitInstallButton) 

   $submitUninstallButton = New-Object System.Windows.Forms.Button 
   $submitUninstallButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitUninstallButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitUninstallButton.Location = New-Object System.Drawing.Size(140,280) 
   $submitUninstallButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitUninstallButton.Text = "Uninstall " 
   $submitUninstallButton.Add_Click({WinGetUninstaller}) 
   $Form.Controls.Add($submitUninstallButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
