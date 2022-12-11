<#
    .SYNOPSIS
        Generates Powershell code to query Windows Event Viewer.

    .DESCRIPTION
        Powershell code is generated using field values on a Windows Form. 
        Search critera is collected and added to a hash table, and then passed 
        to Get-WinEvent.

    .AUTHOR
        clint@clintlang.com
#> 

function main {
<#
    .SYNOPSIS
        Main function used for painting Window Forms GUI.
#>
    Add-Type -AssemblyName System.Windows.Forms;

    ## Collect all available logs
    $logList = (Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Sort-Object RecordCount -Descending).Logname;
    $providerList = Get-WinEvent -ListProvider * -ErrorAction SilentlyContinue;

    ## Dictionary to look up Level enum value
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

    $tooltip = New-Object System.Windows.Forms.ToolTip;
    $tooltip.InitialDelay = 1000;
    $tooltip.AutoPopDelay = 5000;
    $tooltip.ReshowDelay = 500;

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
    $tooltip.SetToolTip($logField, "Log Name");
    $logField.Add_LostFocus({ updateProviders; });
    addCtrl($logField);

    $providerLabel = New-Object System.Windows.Forms.Label;
    $providerLabel.Text = "Provider: ";
    $providerLabel.Width = 100;
    $providerLabel.Location = "10, 50";
    addCtrl($providerLabel);

    $providerField = New-Object System.Windows.Forms.ComboBox;
    $providerField.Width = 260;
    $providerField.Location = "110, 45";
    $tooltip.SetToolTip($providerField, "Provider/Source");
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
    $tooltip.SetToolTip($levelField, "Severity Level");
    addCtrl($levelField);

    $idLabel = New-Object System.Windows.Forms.Label;
    $idLabel.Text = "Event ID: ";
    $idLabel.Width = 100;
    $idLabel.Location = "10, 130";
    addCtrl($idLabel);

    $idField = New-Object System.Windows.Forms.TextBox;
    $idField.Location = "110, 125";
    $tooltip.SetToolTip($idField, "Event ID");
    addCtrl($idField);

    $dateLabel1 = New-Object System.Windows.Forms.Label;
    $dateLabel1.Text = "Start Date: ";
    $dateLabel1.Location = "10, 170";
    $dateLabel1.Width = 100;
    addCtrl($dateLabel1);

    $datePicker1 = New-Object System.Windows.Forms.DateTimePicker;
    $datePicker1.Location = "110, 165";
    $datePicker1.Value = (Get-Date).AddDays("-1").Date;
    $datePicker1.MaxDate = (Get-Date);
    $tooltip.SetToolTip($datePicker1, "Show events after this date");
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
    $tooltip.SetToolTip($datePicker2, "No events after this date");
    addCtrl($datePicker2);

    $maxLabel = New-Object System.Windows.Forms.Label;
    $maxLabel.Text = "Max Events: ";
    $maxLabel.Width = 100;
    $maxLabel.Location = "10, 245";
    addCtrl($maxLabel);

    $maxField = New-Object System.Windows.Forms.TextBox;
    $maxField.Width = 40;
    $maxField.Location = "110, 240";
    $tooltip.SetToolTip($maxField, "Leave blank for no limit");
    addCtrl($maxField);

    $btnSubmit = New-Object System.Windows.Forms.Button;
    $btnSubmit.Location = "10, 280";
    $btnSubmit.Text = "Submit";
    $tooltip.SetToolTip($btnSubmit, "Build Command");
    $btnSubmit.Add_Click({  $rtbResult.Text = buildCmd; });
    addCtrl($btnSubmit);

    $btnCancel = New-Object System.Windows.Forms.Button;
    $btnCancel.Location = "110, 280";
    $btnCancel.Text = "Cancel";
    $tooltip.SetToolTip($btnCancel, "Close this window");
    $btnCancel.Add_Click({ $mainForm.Close(); });
    addCtrl($btnCancel);

    $grpOutpt = New-Object System.Windows.Forms.GroupBox;
    $grpOutpt.Text = "Output: ";
    $grpOutpt.Width = 260;
    $grpOutpt.Height = 50;
    $grpOutpt.Location = "200, 260";
    addCtrl($grpOutpt);


    function chkItm([System.Windows.Forms.CheckBox]$chkbx) {
    <#
        .SYNOPSIS
            Helper function controls check boxes like radio buttons.

        .DESCRIPTION
            Forces only one CheckBox to be selected at a time.
            Unchecks all CheckBoxes in the Control Group, sans the passed in object. 
        
        .PARAMETER $chkbx
            CheckBox Control Object Being Selected 
    #>
        if($chkbx.CheckState -eq [System.Windows.Forms.CheckState]::Checked) {
            foreach ($c in $grpOutpt.Controls) {
                if ($c.Name -ne $chkbx.Name) {
                    $c.Checked = [System.Windows.Forms.CheckState]::Unchecked;
                }
            }
        }
    }

    $outGrid = New-Object System.Windows.Forms.CheckBox;
    $outGrid.Name = "Grid-View";
    $outGrid.Text = $outGrid.Name;
    $outGrid.Width = "75";
    $outGrid.Location = "10, 20";
    $outGrid.Add_CheckStateChanged({ chkItm($outGrid); });
    $tooltip.SetToolTip($outGrid, "This output method requires an active desktop session.");
    $grpOutpt.Controls.Add($outGrid);

    $outHtml = New-Object System.Windows.Forms.CheckBox;
    $outHtml.Name = "HTML";
    $outHtml.Text = $outHtml.Name;
    $outHtml.Location = "100, 20";
    $outHtml.Width = "75";
    $outHtml.Add_CheckStateChanged({ chkItm($outHtml); });
    $tooltip.SetToolTip($outHtml, "Can be used for active desktop or console session.");
    $grpOutpt.Controls.Add($outHtml);

    $outJson = New-Object System.Windows.Forms.CheckBox;
    $outJson.Name = "JSON";
    $outJson.Text = $outJson.Name;
    $outJson.Location = "180, 20";
    $outJson.Width = "75";
    $outJson.Add_CheckStateChanged({ chkItm($outJson); });
    $tooltip.SetToolTip($outJson, "Useful for exporting objects.");
    $grpOutpt.Controls.Add($outJson);   

    function svDlg([System.IO.FileInfo]$fn) {
    <#
        .SYNOPSIS
            Helper function for SaveFileDialog

        .DESCRIPTION
            This SaveFileDialog is a multi-purpose Control.
            File name is passed into parameter 0 and the
            extension is evaluated by a Switch to determine
            function behaviour.

        .PARAMETER fn 
            Specifies the current or default file name.
            File name is stored as FileInfo for versatility.
    #>

        switch($fn.Extension)
        {
            { [Regex]::IsMatch($PSItem, "^.html?$")  } 
                    { 
                        $target = "HtmlFile";
                        $filter = "HTML Files (*.htm, *.html)|*.html;*.htm";
                    }

            { [Regex]::IsMatch($PSItem, "^.json$")  }
                    { 
                        $target = "JsonFile";
                        $filter = "JSON Files (*.json) | *.json;";
                    }

            default { return; }
        }

        $svDlg = New-Object System.Windows.Forms.SaveFileDialog;
        $svDlg.FileName = $fn.Name;
        $svDlg.InitialDirectory = $fn.Directory;
        $svDlg.Filter = $filter;
        $svDlg.Add_FileOk({ Invoke-Expression -Command "`$Script:$target = `"$($svDlg.FileName)`""; });
        $svDlg.ShowDialog();
    }

    $outptMenu = New-Object System.Windows.Forms.ContextMenu;

    ## Set HTML File
    $setHtmlFN = $outptMenu.MenuItems.Add("Set HTML File Name");
    $Script:HtmlFile = "C:\temp\WinEvents-output.html";
    $setHtmlFN.Add_Click({ svDlg($Script:HtmlFile); });

    ## Option to open HTML after execution. 
    $Script:Htmlopen = $false;
    $setHtmlopen = $outptMenu.MenuItems.Add("Open HTML File");
    $setHtmlopen.Add_Click({
        ## Switch ContextMenu item
        switch($Script:Htmlopen)
        {
            $true
            {
                ## Setting switched off
                $setHtmlopen.Checked = $false;
                $Script:Htmlopen = $false;
            }
            $false
            {
                ## Setting switched on
                $setHtmlopen.Checked = $true;
                $Script:Htmlopen = $true;
            }
        }
    });

    ## Set JSON File
    $selJson = $outptMenu.MenuItems.Add("Set JSON File Name");
    $script:JsonFile = "C:\temp\WinEvents-output.json";
    $selJson.Add_Click({ svDlg($Script:JsonFile); });

    $grpOutpt.ContextMenu = $outptMenu;

    $rtbResult = New-Object System.Windows.Forms.RichTextBox;
    $rtbResult.Width = 475;
    $rtbResult.Height = 150;
    $rtbResult.Location = "10, 320";
    addCtrl($rtbResult);

    $btnRun = New-Object System.Windows.Forms.Button;
    $btnRun.Location = "10, 480";
    $btnRun.Text = "Run";
    $tooltip.SetToolTip($btnRun, "Execute code in results box");
    $btnRun.Add_Click({ runCMD; });
    addCtrl($btnRun);

    $btnCopy = New-Object System.Windows.Forms.Button;
    $btnCopy.Location = "110, 480";
    $btnCopy.Text = "Copy";
    $tooltip.SetToolTip($btnCopy, "Copy code in results box to clipboard");
    $btnCopy.Add_Click({ Set-Clipboard $rtbResult.Text; });
    addCtrl($btnCopy);

    $mainForm.ShowDialog() | Out-Null;
}

function addCtrl ($type) {
<#
    .SYNOPSIS
        Helper function used to add Controls to the Main Form.
#> 
    try {
        $mainForm.Controls.Add($type);
    }
    catch {
        newError "Failed to add control $($type.GetType().Name)";
    }
}

function updateProviders {
<#
    .SYNOPSIS
        Update list of source providers.

    .DESCRIPTION
        Query source providers for the selected log, then add to ComboBox items.
        ComboBox list items are updated when the logname changes.
#>
    try {
        $providerField.Items.Clear() | Out-Null;
        $providerField.Text = [System.String]::Empty | Out-Null;
        $providerList | Where-Object {$_.LogLinks.LogName -EQ $logField.Text} | ForEach-Object {$providerField.Items.Add($_.Name);} | Out-Null;
    }
    catch {
        newError "Failed to update provider";
    }
}

function newError ($msg) {
<#
    .SYNOPSIS
        Helper function for exception handling.

    .DESCRIPTION
        When an exception is caught a MessageBox style error message is displayed.
#>
        [System.Windows.Forms.MessageBox]::Show("$msg`r`n$($_.Exception.Message)","Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
}

function notNull ($str) {
<#
    .SYNOPSIS
        Helper function determines if attribute is set.
#>
    if(![System.String]::IsNullOrEmpty($str)) { return $true; }
    else { return $false }
}

function buildCmd {
<#
    .SYNOPSIS
        Build Powershell command from Form input
#>
    ## Get input from fields in Form
    $log = $logField.Text;
    $provider = $providerField.Text;
    $dateStart = $datePicker1.Value;
    $dateEnd = $datePicker2.Value;
    $level = $levelField.Text;
    $id = $idField.Text;
    $max = $maxField.Text;

    ## Look for selected output option
    $checked = [System.Windows.Forms.CheckState]::Checked;
    $grid = $outGrid.CheckState -eq $checked;
    $html = $outHtml.CheckState -eq $checked;
    $json = $outJson.CheckState -eq $checked;

    $attr = @();
    ## If attribute exists add to list.  
    if (notNull($log))        { $attr += "LogName = `"$log`";"; }
    if (notNull($provider))   { $attr += "ProviderName = `"$provider`";"; }
    if (notNull($dateStart))  { $attr += "StartTime = `"$dateStart`";"; }
    if (notNull($dateEnd))    { $attr += "EndTime = `"$dateEnd`";"; }
    if (notNull($level))      { $attr += [System.String]::Format("Level = `"{0}`";", $lvlTypes.Item($level)); }
    if (notNull($id))         { $attr += "ID = $id;"; }

    ## Start building command
    $cmd = "Get-WinEvent -FilterHashtable @{ $attr }";

    ## Add MaxEvents property if specified 
    if (notNull($max)) { $cmd += " -MaxEvents $max"; }

    ## Control output options
    switch ($true) {
        $grid 
        {
            ## Send object data to Powershell Grid-View
            $cmd += " | Select TimeCreated,ID,LevelDisplayName,ProviderName,Message | Out-GridView -Title `"WinEvents - $log`" -PassThru;";       
        }

        $html
        {
            ## Send object data to HTML Table
            $css = buildCss;
            #$fn = [System.IO.Path]::GetTempFileName();
            #$fn = $fn.Replace(".tmp", ".html");
            $cmd = "`$events = $cmd;";` 
            $cmd += "`$events| Select TimeCreated,ID,LevelDisplayName,ProviderName,Message | ConvertTo-Html -Title `"WinEvents - $log`" -CssUri `"$css`" | Out-File `"$HtmlFile`";";
            if ($Htmlopen) { $cmd += "Invoke-Item `"$HtmlFile`";"; }
        }

        $json
        {
            ## Convert object data to JSON for file transfer
            $cmd = "`$events = $cmd;";
            $cmd += "`$events| Select TimeCreated,ID,LevelDisplayName,ProviderName,Message | ConvertTo-Json | Out-File `"$JsonFile`";";
        }
    }

    return $cmd;
}

function buildCss {
<#
    .SYNOPSIS
        Save CSS file for HTML Table Output mode

    .DESCRIPTION
        CSS file is saved in the environemt temp location.
        Function returns path to CSS file.             
#>
    
    ## CSS
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

    ## Write CSS file and return filename
    $cssFile = Join-Path ([System.IO.Path]::GetTempPath()) "WinEvent-styles.css";
    $css | Out-File $cssFile;
    return $cssFile;
}

function runCMD {
<#
    .SYNOPSIS
        Execute command in results text box.

    .DESCRIPTION
        Function executes when the Run button is selected.
        Any code in the results TextBox is saved to a temp file in %TEMP%.
        Script file is passed into a new Powershell process.
#>
    ## Get code from TextBox
    $cmd = $rtbResult.Text;

    ## Check for value
    if($cmd.Length -lt 1) { 
        newError "Command text is empty.";
        return;
     }

     ## Add error handling 
     $cmd = "
     `$ErrorActionPreference = 'Stop';
     Add-Type -AssemblyName System.Windows.Forms;
     try {
            $cmd
     }
     catch {
            `$e = `$_.Exception.Message;
            [System.Windows.Forms.MessageBox]::Show(`"(`$e)`", `"Doh!`", `"OK`", `"Error`");  
     }";

     ## Save code ito temp file 
    $outfile = Join-Path ([System.IO.Path]::GetTempPath()) "WinEvents_output.ps1";
    $cmd | Out-File $outfile

    ## Start Process
    try {
        [System.IO.FileInfo]$ps =  Join-Path ([System.Environment]::SystemDirectory) "WindowsPowerShell\v1.0\powershell.exe";
        Start-Process $ps -ArgumentList "-NoLogo -File $outfile" | Out-Null;
    }
    catch {
        newError "Powershell failed to start.";
    }
}

##

main;
