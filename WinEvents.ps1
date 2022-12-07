Add-Type -AssemblyName System.Windows.Forms;

$logList = (Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Sort-Object RecordCount -Descending).Logname

$lvlTypes = @{};
$lvlTypes.Add("Critical", 1);
$lvlTypes.Add("Error", 2);
$lvlTypes.Add("Warning", 3);
$lvlTypes.Add("Information", 4);

function addCtrl ($type) { $mainForm.Controls.Add($type); }

function updateProviders {
    $providerList = Get-WinEvent -ListProvider * -ErrorAction SilentlyContinue;
    $providerList | Where-Object {$_.LogLinks.LogName -EQ $logField.Text} | ForEach-Object {$providerField.Items.Add($_.Name);} | Out-Null
}

$mainForm = New-Object System.Windows.Forms.Form;
$mainForm.Width = 500;
$mainForm.Height = 550;
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle;

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
$idField.Width = 200;
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
addCtrl($datePicker1);

$dateLabel2 = New-Object System.Windows.Forms.Label;
$dateLabel2.Text = "End Date: ";
$dateLabel2.Location = "10, 210";
$dateLabel2.Width = 100;
addCtrl($dateLabel2);

$datePicker2 = New-Object System.Windows.Forms.DateTimePicker;
$datePicker2.Location = "110, 205";
$datePicker2.Value = (Get-Date);
$mainForm.Controls.Add($datePicker2);

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
$btnSubmit.Add_MouseDown({  $rtbResult.Text = buildCmd; });
addCtrl($btnSubmit);

$btnCancel = New-Object System.Windows.Forms.Button;
$btnCancel.Location = "110, 280";
$btnCancel.Text = "Cancel";
$btnCancel.Add_MouseDown({ $mainForm.Close(); });
addCtrl($btnCancel);

$rtbResult = New-Object System.Windows.Forms.RichTextBox;
$rtbResult.Width = 475;
$rtbResult.Height = 150;
$rtbResult.Location = "10, 320";
addCtrl($rtbResult);

$btnRun = New-Object System.Windows.Forms.Button;
$btnRun.Location = "10, 490";
$btnRun.Text = "Run";
addCtrl($btnRun);

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

    $attr = @();
    if (notNull($log)) { $attr += "LogName = `"$log`";"; }
    if (notNull($providerField)) { $attr += "ProviderName = `"$provider`";"; }
    if (notNull($dateStart)) { $attr += "StartTime = `"$dateStart`";"; }
    if (notNull($dateEnd)) { $attr += "EndTime = `"$dateEnd`";"; }
    if (notNull($level)) { 
        $lvlEnum = $lvlTypes.Item($level);
        $attr += "Level = `"$lvlEnum`"";
     }
    if (notNull($id)) { $attr += "ID = $id;"; }

    $cmd = "Get-WinEvent -FilterHashtable @{ $attr }";
    if (notNull($max)) { $cmd += " -MaxEvents $max"; }

    return $cmd;
}

##
$mainForm.ShowDialog() | Out-Null;

<#
if ($btnSubmit.DialogResult -eq "OK") {
    $mainForm.Controls.Clear();
    addCtrl($rtbResult);
    $mainForm.ShowDialog();
    
    try {
        Invoke-Expression (buildCmd);
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow -BackgroundColor Black;
    }
}
#>
