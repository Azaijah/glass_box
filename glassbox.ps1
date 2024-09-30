#Install-Module -Name ImportExcel -Scope CurrentUser

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(510,295)
$Form.text                       = "Glass Box"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")
$Form.minimumSize                = New-Object System.Drawing.Size(500,300)
$Form.maximumSize                = New-Object System.Drawing.Size(500,300)  

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "IP: "
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(11,15)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



#IP RANGE TEXT BOXES #####################################
$rangebox1_1                        = New-Object system.Windows.Forms.TextBox
$rangebox1_1.Visible                = $false
$rangebox1_1.multiline              = $false
$rangebox1_1.width                  = 30
$rangebox1_1.height                 = 20
$rangebox1_1.location               = New-Object System.Drawing.Point(10,46)
$rangebox1_1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox1_2                        = New-Object system.Windows.Forms.TextBox
$rangebox1_2.Visible                = $false
$rangebox1_2.multiline              = $false
$rangebox1_2.width                  = 30
$rangebox1_2.height                 = 20
$rangebox1_2.location               = New-Object System.Drawing.Point(40,46)
$rangebox1_2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox1_3                        = New-Object system.Windows.Forms.TextBox
$rangebox1_3.Visible                = $false
$rangebox1_3.multiline              = $false
$rangebox1_3.width                  = 30
$rangebox1_3.height                 = 20
$rangebox1_3.location               = New-Object System.Drawing.Point(70,46)
$rangebox1_3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox1_4                      = New-Object system.Windows.Forms.TextBox
$rangebox1_4.Visible                = $false
$rangebox1_4.multiline              = $false
$rangebox1_4.width                  = 30
$rangebox1_4.height                 = 20
$rangebox1_4.location               = New-Object System.Drawing.Point(100,46)
$rangebox1_4.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox2_1                        = New-Object system.Windows.Forms.TextBox
$rangebox2_1.Visible                = $false
$rangebox2_1.multiline              = $false
$rangebox2_1.width                  = 30
$rangebox2_1.height                 = 20
$rangebox2_1.location               = New-Object System.Drawing.Point(150,46)
$rangebox2_1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox2_2                        = New-Object system.Windows.Forms.TextBox
$rangebox2_2.Visible                = $false
$rangebox2_2.multiline              = $false
$rangebox2_2.width                  = 30
$rangebox2_2.height                 = 20
$rangebox2_2.location               = New-Object System.Drawing.Point(180,46)
$rangebox2_2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox2_3                        = New-Object system.Windows.Forms.TextBox
$rangebox2_3.Visible                = $false
$rangebox2_3.multiline              = $false
$rangebox2_3.width                  = 30
$rangebox2_3.height                 = 20
$rangebox2_3.location               = New-Object System.Drawing.Point(210,46)
$rangebox2_3.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$rangebox2_4                      = New-Object system.Windows.Forms.TextBox
$rangebox2_4.Visible                = $false
$rangebox2_4.multiline              = $false
$rangebox2_4.width                  = 30
$rangebox2_4.height                 = 20
$rangebox2_4.location               = New-Object System.Drawing.Point(240,46)
$rangebox2_4.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$range_label                          = New-Object system.Windows.Forms.Label
$range_label.Visible                  = $false
$range_label.text                     = "-"
$range_label.AutoSize                 = $true
$range_label.width                    = 25
$range_label.height                   = 10
$range_label.location                 = New-Object System.Drawing.Point(135,50)
$range_label.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



#####################################

#IP INDIVIDUAL 

#IP RANGE TEXT BOXES #####################################
$invidbox1                        = New-Object system.Windows.Forms.TextBox
$invidbox1.multiline              = $false
$invidbox1.Visible                = $true
$invidbox1.width                  = 125
$invidbox1.height                 = 20
$invidbox1.location               = New-Object System.Drawing.Point(11,46)
$invidbox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#####################################

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(400,210)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button1.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#0076ff")

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "Individual"
$ComboBox1.width                 = 100
$ComboBox1.height                = 20
@('Individual','Range') | ForEach-Object {[void] $ComboBox1.Items.Add($_)}
$ComboBox1.location              = New-Object System.Drawing.Point(35,8)
$ComboBox1.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($Label1,$rangebox1_1,$rangebox1_2,$rangebox1_3,$rangebox1_4,$rangebox2_1,$rangebox2_2,$rangebox2_3,$rangebox2_4,$Button1,$range_label,$ComboBox1,$invidbox1))

$Button_Click = 
{

    if($ComboBox1.Text -eq "Individual") {
        $Inputstring =  $invidbox1.Text
        $hosts = $InputString.Split(",")
    } else {

        $hosts=$rangebox1_4.Text..$rangebox2_4.Text | %{"10.100.10.$_"}
    }


    $add = 2
    
    Foreach ($h in $hosts)
    {

        $hostname = ([System.Net.Dns]::GetHostByAddress($h)).HostName
        $app_info = get-wmiobject Win32_Product -computername $hostname | Select-Object Name,InstallSource,InstallDate,Version #| Export-Excel -Path ./$hostname.xlsx 
        $os_info = get-wmiobject Win32_OperatingSystem -computername $hostname | Select-Object Name,OSArchitecture,InstallDate,Version  #| Export-Excel -Path ./$hostname.xlsx 
        $patch_info = get-wmiobject Win32_QuickFixEngineering -computername $hostname | Select-Object Description,HotFixID,InstalledOn  #| Export-Excel -Path ./$hostname.xlsx
        $net_info =  get-wmiobject  Win32_NetworkAdapter -computername $hostname | Select-Object ServiceName,MACAddress,AdapterType,DeviceID 
        $user_info = get-wmiobject  Win32_UserAccount -computername $hostname | Select-Object Name

        $add = $add - 1
        $hostname | Export-Excel -StartRow $add ./inforeport.xlsx 
        $add = $add + 2
        $os_info | Export-Excel -AutoSize -TableName os_info -StartRow $add -Path ./inforeport.xlsx
        $add = $add + $os_info.length + 3
        $net_info | Export-Excel -AutoSize -TableName net_info -StartRow $add -Path ./inforeport.xlsx
        $add = $add + $net_info.length + 2
        $user_info | Export-Excel -AutoSize -TableName user_info -StartRow $add -Path ./inforeport.xlsx
        $add = $add + $user_info.length + 2
        $patch_info | Export-Excel -AutoSize -TableName patch_info -StartRow $add -Path ./inforeport.xlsx
        $add = $add + $patch_info.length + 2
        $app_info | Export-Excel -AutoSize -TableName app_info -StartRow $add ./inforeport.xlsx
        $add = $add + $app_info.length + 4
 
    }

    $Form.Close()
}


$BoxSelection = 
{
    if($ComboBox1.Text -eq "Individual") {
        $rangebox1_1.Visible                = $false
        $rangebox1_2.Visible                = $false
        $rangebox1_3.Visible                = $false
        $rangebox1_4.Visible                = $false

        $rangebox2_1.Visible                = $false
        $rangebox2_2.Visible                = $false
        $rangebox2_3.Visible                = $false
        $rangebox2_4.Visible                = $false

        $range_label.Visible                  = $false
        $invidbox1.Visible                = $true
    } else {
        $rangebox1_1.Visible                = $true
        $rangebox1_2.Visible                = $true
        $rangebox1_3.Visible                = $true
        $rangebox1_4.Visible                = $true

        $rangebox2_1.Visible                = $true
        $rangebox2_2.Visible                = $true
        $rangebox2_3.Visible                = $true
        $rangebox2_4.Visible                = $true

        $range_label.Visible                   = $true
        $invidbox1.Visible                = $false
    }
    

}

$ComboBox1.add_SelectedIndexChanged($BoxSelection)

$Button1.Add_Click($Button_Click)

[void]$Form.ShowDialog()