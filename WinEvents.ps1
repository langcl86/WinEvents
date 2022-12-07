function main {}
    Add-Type -AssemblyName System.Windows.Forms;

    $logList = (Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Sort-Object RecordCount -Descending).Logname

    $lvlTypes = @{};
    $lvlTypes.Add("Critical", 1);
    $lvlTypes.Add("Error", 2);
    $lvlTypes.Add("Warning", 3);
    $lvlTypes.Add("Information", 4);

    $mainForm = New-Object System.Windows.Forms.Form;
    $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle;
    $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $mainForm.Width = 500;
    $mainForm.Height = 550;
    $mainForm.Text = "WinEvents";

    $logLabel = New-Object System.Windows.Forms.Label;
    $logLabel.Text = "Log Name: ";
    $logLabel.Location = "10, 10";
    $logLabel.Width = 100;
    addCtrl($logLabel);

    $logField = New-Object System.Windows.Forms.ComboBox;
    $logList | Foreach-Object { $logField.Items.Add($_) | Out-Null; }
    $logField.SelectedItem = "System";
    $logField.Location = "110, 5";
    $logField.Width = 260;
    $logField.Add_SelectedValueChanged({ updateProviders; });
    addCtrl($logField);

    $providerLabel = New-Object System.Windows.Forms.Label;
    $providerLabel.Text = "Provider: ";
    $providerLabel.Width = 100;
    $providerLabel.Location = "10, 50";
    addCtrl($providerLabel);

    $providerField = New-Object System.Windows.Forms.ComboBox;
    $providerField.Width = 260;
    $providerField.Location = "110, 45";
    updateProviders;
    addCtrl($providerField);

    $levelLabel = New-Object System.Windows.Forms.Label;
    $levelLabel.Text = "Level: ";
    $levelLabel.Width = 100;
    $levelLabel.Location = "10, 95"
    addCtrl($levelLabel);

    $levelField = New-Object System.Windows.Forms.ComboBox;
    $lvlTypes.Keys | ForEach-Object { $levelField.Items.Add($_); } | Out-Null;
    $levelField.Width = 200;
    $levelField.Location = "110, 85";
    addCtrl($levelField);

    $idLabel = New-Object System.Windows.Forms.Label;
    $idLabel.Text = "Event ID: ";
    $idLabel.Width = 100;
    $idLabel.Location = "10, 130";
    addCtrl($idLabel);

    $idField = New-Object System.Windows.Forms.TextBox;
    $idField.Location = "110, 125";
    addCtrl($idField);

    $dateLabel1 = New-Object System.Windows.Forms.Label;
    $dateLabel1.Text = "Start Date: ";
    $dateLabel1.Location = "10, 170";
    $dateLabel1.Width = 100;
    addCtrl($dateLabel1);

    $datePicker1 = New-Object System.Windows.Forms.DateTimePicker;
    $datePicker1.Location = "110, 165";
    $datePicker1.Value = (Get-Date).AddDays("-1");
    $datePicker1.MaxDate = (Get-Date);
    $datePicker1.Add_ValueChanged({$datePicker2.MinDate = $datePicker1.Value;});
    addCtrl($datePicker1);

    $dateLabel2 = New-Object System.Windows.Forms.Label;
    $dateLabel2.Text = "End Date: ";
    $dateLabel2.Location = "10, 210";
    $dateLabel2.Width = 100;
    addCtrl($dateLabel2);

    $datePicker2 = New-Object System.Windows.Forms.DateTimePicker;
    $datePicker2.Width = 200;
    $datePicker2.Location = "110, 205";
    $datePicker2.Value = (Get-Date);
    $datePicker2.MinDate = $datePicker1.Value;
    $datePicker2.MaxDate = (Get-Date);
    addCtrl($datePicker2);

    $maxLabel = New-Object System.Windows.Forms.Label;
    $maxLabel.Text = "Max Events: ";
    $maxLabel.Width = 100;
    $maxLabel.Location = "10, 245";
    addCtrl($maxLabel);

    $maxField = New-Object System.Windows.Forms.TextBox;
    $maxField.Width = 40;
    $maxField.Location = "110, 240";
    $maxField.Text = "10";
    addCtrl($maxField);

    $btnSubmit = New-Object System.Windows.Forms.Button;
    $btnSubmit.Location = "10, 280";
    $btnSubmit.Text = "Submit";
    $btnSubmit.Add_Click({  $rtbResult.Text = buildCmd; });
    addCtrl($btnSubmit);

    $btnCancel = New-Object System.Windows.Forms.Button;
    $btnCancel.Location = "110, 280";
    $btnCancel.Text = "Cancel";
    $btnCancel.Add_Click({ $mainForm.Close(); });
    addCtrl($btnCancel);

    $grpOptns = New-Object System.Windows.Forms.GroupBox;
    $grpOptns.Text = "Output: ";
    $grpOptns.Width = 260;
    $grpOptns.Height = 50;
    $grpOptns.Location = "200, 260";
    addCtrl($grpOptns);


    function chkItm([System.Windows.Forms.CheckBox]$chkbx) {
        if($chkbx.CheckState -eq [System.Windows.Forms.CheckState]::Checked) {
            foreach ($c in $grpOptns.Controls) {
                if ($c.Name -ne $chkbx.Name) {
                    $c.Checked = [System.Windows.Forms.CheckState]::Unchecked;
                }
            }
            }
    }

    $outGrid = New-Object System.Windows.Forms.CheckBox;
    $outGrid.Name = "Out Grid-View";
    $outGrid.Text = $outGrid.Name;
    $outGrid.Location = "10, 20";
    $outGrid.Add_CheckStateChanged({ chkItm($outGrid); });
    $grpOptns.Controls.Add($outGrid);

    $outHtml = New-Object System.Windows.Forms.CheckBox;
    $outHtml.Name = "Out HTML";
    $outHtml.Text = $outHtml.Name;
    $outHtml.Location = "130, 20";
    $outHtml.Add_CheckStateChanged({ chkItm($outHtml); });
    $grpOptns.Controls.Add($outHtml);

    $svDlg = New-Object System.Windows.Forms.SaveFileDialog;
    $svDlg.FileName = "C:\temp\WinEvents-output.html";
    $svDlg.Filter = "HTML Files (*.htm, *.html)|*.html;*.htm";
    $svDlg.Add_FileOk({ <# Set Filename #> });

    $ctmnu = New-Object System.Windows.Forms.ContextMenu;
    $setHtmlFN = $ctmnu.MenuItems.Add("Set HTML Filename");
    $setHtmlFN.Add_Click({ $svDlg.ShowDialog(); });
    $grpOptns.ContextMenu = $ctmnu;

    $rtbResult = New-Object System.Windows.Forms.RichTextBox;
    $rtbResult.Width = 475;
    $rtbResult.Height = 150;
    $rtbResult.Location = "10, 320";
    addCtrl($rtbResult);

    $btnRun = New-Object System.Windows.Forms.Button;
    $btnRun.Location = "10, 480";
    $btnRun.Text = "Run";
    $btnRun.Add_Click({ runCMD; });

    addCtrl($btnRun);

    $btnCopy = New-Object System.Windows.Forms.Button;
    $btnCopy.Location = "110, 480";
    $btnCopy.Text = "Copy";
    $btnCopy.Add_Click({ Set-Clipboard $rtbResult.Text; });
    addCtrl($btnCopy);

    $mainForm.ShowDialog() | Out-Null;
#}

function addCtrl ($type) { 
    try {
        $mainForm.Controls.Add($type);
    }
    catch {
        newError "Failed to add control $($type.GetType().Name)";
    }
}

function updateProviders {
    try {
        $providerList = Get-WinEvent -ListProvider * -ErrorAction SilentlyContinue;
        $providerList | Where-Object {$_.LogLinks.LogName -EQ $logField.Text} | ForEach-Object {$providerField.Items.Add($_.Name);} | Out-Null;
    }
    catch {
        newError "Failed to update provider";
    }
}

function newError ($msg) {
        [System.Windows.Forms.MessageBox]::Show("$msg`r`n$($_.Exception.Message)","Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
}

function notNull ($str) {
    if(![System.String]::IsNullOrEmpty($str)) { return $true; }
    else { return $false }
}

function buildCmd {
    $log = $logField.Text;
    $provider = $providerField.Text;
    $dateStart = $datePicker1.Value;
    $dateEnd = $datePicker2.Value;
    $level = $levelField.Text;
    $id = $idField.Text;
    $max = $maxField.Text;
    $grid = $outGrid.CheckState -eq [System.Windows.Forms.CheckState]::Checked;
    $html = $outHtml.CheckState -eq [System.Windows.Forms.CheckState]::Checked;

    $attr = @();
    if (notNull($log)) { $attr += "LogName = `"$log`";"; }
    if (notNull($provider)) { $attr += "ProviderName = `"$provider`";"; }
    if (notNull($dateStart)) { $attr += "StartTime = `"$dateStart`";"; }
    if (notNull($dateEnd)) { $attr += "EndTime = `"$dateEnd`";"; }
    if (notNull($level)) { 
        $lvlEnum = $lvlTypes.Item($level);
        $attr += "Level = `"$lvlEnum`"";
     }
    if (notNull($id)) { $attr += "ID = $id;"; }

    $cmd = "Get-WinEvent -FilterHashtable @{ $attr }";
    if (notNull($max)) { $cmd += " -MaxEvents $max"; }

    switch ($true) {
        $grid 
        {
            $cmd += " | Out-GridView -Wait";       
        }

        $html
        {
            $css = buildCss;
            $fn = [System.IO.Path]::GetTempFileName();
            $fn = $fn.Replace(".tmp", ".html");
            $cmd = "`$events = $cmd;";
            $cmd += "`$events| Select TimeCreated,ID,LevelDisplayName,ProviderName,Message | ConvertTo-Html -Title `"WinEvents - $log`" -CssUri `"$css`" | Out-File `"$fn`"; Invoke-Item `"$fn`";";
        }
    }

    return $cmd;
}

function buildCss {
    $cssFile = Join-Path ([System.IO.Path]::GetTempPath()) "WinEvent-styles.css";
    $css = "
    table {
       width: 98%;
       background-color: #000000;
       border: 3px solid #000000;
       border-spacing: 0px;
    }

    th, td {
        border: 1px solid #000000;
    }

    tbody tr:nth-child(odd){
      background-color: #eeeeee;
      color: #000000;
    }
    tbody tr:nth-child(even){
      background-color: #dddddd;
      color: #000000;
    }

    th {
        Background-Color: #efefef;
        Color: #000000;
        font-family: Verdana;
        padding: 5px;
    }
    ";

    $css | Out-File $cssFile;
    return $cssFile;
}

function runCMD {
    $cmd = $rtbResult.Text;

    if($cmd.Length -lt 1) { 
        newError "Command text is empty.";
        return;
     }
        
    [System.IO.FileInfo]$ps =  Join-Path ([System.Environment]::SystemDirectory) "WindowsPowerShell\v1.0\powershell.exe";
    $outfile = Join-Path ([System.IO.Path]::GetTempPath()) "WinEvents_output.ps1";
    $cmd | Out-File $outfile

    try {
        Start-Process $ps -ArgumentList "-NoLogo -File $outfile" | Out-Null;
    }
    catch {
        newError "Powershell failed to start.";
    }
}

##

main;