[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

####Disable ALT+F4 from closing form#####
$form_KeyDown=[System.Windows.Forms.KeyEventHandler]{
    #Event Argument: $_ = [System.Windows.Forms.KeyEventArgs]

    if ($_.Alt -eq $true -and $_.KeyCode -eq 'F4') {
        $script:altF4Pressed = $true;           
    }
}

$form_FormClosing=[System.Windows.Forms.FormClosingEventHandler]{
    #Event Argument: $_ = [System.Windows.Forms.FormClosingEventArgs]

    if ($script:altF4Pressed)
    {
        if ($_.CloseReason -eq 'UserClosing') {
            $_.Cancel = $true
            $script:altF4Pressed = $false;
        }
    }
}



$Font1 = New-Object System.Drawing.Font("Century Gothic",16)
$Font1Bold = New-Object System.Drawing.Font("Century Gothic",16,[System.Drawing.FontStyle]::Bold)

####Build Main Form####
$Form = New-Object System.Windows.Forms.Form
$Form.width = 900
$Form.height = 350
#$Form.ForeColor = $colors[$colorchoice]
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$Form.Text = "Your Windows 10 is no longer supported..."
$Form.maximumsize = New-Object System.Drawing.Size(1000,1000)
$Form.startposition = "centerscreen"
$Form.KeyPreview = $True
$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter") {}})
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape") {Write-Host "Delay"
$form.close()}})
$Form.Font = $Font1
#$Form.AutoSize = $True
#$Form.AutoSizeMode = "GrowAndShrink"
$Form.ControlBox = $False
$Form.KeyPreview = $True
$Form.add_KeyDown($form_keydown)
$Form.add_FormClosing($form_FormClosing)
$Form.topmost = $True

####Build Upgrade Clicked Form####
$upgradeform = New-Object System.Windows.Forms.Form
$upgradeform.width = 900
$upgradeform.height = 350
$upgradeform.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$upgradeform.Text = "Upgrade Running"
$upgradeform.maximumsize = New-Object System.Drawing.Size(1000,1000)
$upgradeform.startposition = "centerscreen"
$upgradeform.KeyPreview = $True
$upgradeform.Add_KeyDown({if ($_.KeyCode -eq "Enter") {}})
$upgradeform.Add_KeyDown({if ($_.KeyCode -eq "Escape") {
$upgradeform.close()}})
$upgradeform.Font = $Font1
$upgradeform.ControlBox = $False
$upgradeform.KeyPreview = $True
$upgradeform.add_KeyDown($upgradeform_keydown)
$upgradeform.add_FormClosing($upgradeform_FormClosing)
$UpgradeForm.topmost = $True

####Import CIT Logo for Form####
$img = [System.Drawing.Image]::Fromfile("C:\Windows\LTSvc\download\citlogo.png")
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width = 116
$pictureBox.Height = 117
$pictureBox.SizeMode = 'StretchImage'
$pictureBox.Image = $img
$pictureBox.Location = new-object System.Drawing.Size((($form.Width-$pictureBox.Width) / 2),0)
$form.controls.add($pictureBox)

####Import CIT Logo for upgradeform####
$img = [System.Drawing.Image]::Fromfile("C:\Windows\LTSvc\download\citlogo.png")
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width = 116
$pictureBox.Height = 117
$pictureBox.SizeMode = 'StretchImage'
$pictureBox.Image = $img
$pictureBox.Location = new-object System.Drawing.Size((($upgradeform.Width-$pictureBox.Width) / 2),0)
$upgradeform.controls.add($pictureBox)

####Build main text for Form####
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Width = $form.Width
$Label1.Height = 130
$Label1.Font = $Font1Bold
$Label1.Text = "You are running Windows 10 version @buildnumber@ which is no longer supported`r`nYou must upgrade to the latest version as soon as possible`r`nFor more information please speak to the CIT helpdesk or your Technical Consultant"
$Label1.TextAlign="TopCenter"
$Label1.Location = new-object System.Drawing.Size(0,($pictureBox.Bottom+10))
$label1.BackColor = [System.Drawing.Color]::FromName("Transparent")
$Form.Controls.Add($Label1)

####Build main text for Upgrade Form####
$UpgradeLabel1 = New-Object System.Windows.Forms.Label
$UpgradeLabel1.Width = $form.Width
$UpgradeLabel1.Height = 100
$UpgradeLabel1.Font = $Font1Bold
$UpgradeLabel1.Text = "Upgrade files will now be downloaded`r`nThe time this takes varies depending on your connection speed`r`nYou will be prompted again when the upgrade is ready to start"
$UpgradeLabel1.TextAlign="TopCenter"
$UpgradeLabel1.Location = new-object System.Drawing.Size(0,($pictureBox.Bottom+10))
$UpgradeLabel1.BackColor = [System.Drawing.Color]::FromName("Transparent")
$UpgradeForm.Controls.Add($UpgradeLabel1)

####Build Delay Button####
$Button1 = new-object System.Windows.Forms.Button
$Button1.Location = new-object System.Drawing.Size(20,($Label1.Bottom))
$Button1.AutoSize = $True
$Button1.Font = $Font1
$Button1.Text = "Delay for 1 day"
$Button1.ForeColor = "ControlText"
$Button1.BackColor = "Control"
$Button1.Add_Click({
Write-Host "Delay"
$form.close()
})

####Build Upgrade Button####
$Button2 = new-object System.Windows.Forms.Button
$Button2.Location = new-object System.Drawing.Size(($label1.parent.Width-200),($Label1.Bottom))
$Button2.AutoSize = $True
$Button2.Font = $Font2
$Button2.Text = "Upgrade Now"
$Button2.ForeColor = "ControlText"
$Button2.BackColor = "Control"
$Button2.Add_Click({
Write-Host "Upgrade"
$Upgradeform.Controls.Add($UpgradeOKButton)
$UpgradeForm.Add_Shown({$UpgradeForm.Activate()})
$upgradeform.ShowDialog()
})

####Build Upgrade OK Button####
$UpgradeOKButton = new-object System.Windows.Forms.Button
$UpgradeOKButton.Location = new-object System.Drawing.Size((($Upgradelabel1.parent.Width-$UpgradeOkButton.Width)/2),($UpgradeLabel1.Bottom))
$UpgradeOKButton.AutoSize = $True
$UpgradeOKButton.Font = $Font2
$UpgradeOKButton.Text = "OK"
$UpgradeOKButton.ForeColor = "ControlText"
$UpgradeOKButton.BackColor = "Control"
$UpgradeOKButton.Add_Click({
$Upgradeform.close() 
$form.close()
})


####Add controls and activate####
$Form.Controls.Add($Button2)
$Form.Controls.Add($Button1)
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog() | Out-Null

