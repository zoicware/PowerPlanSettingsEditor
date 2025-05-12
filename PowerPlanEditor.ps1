If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
  Start-Process PowerShell.exe -ArgumentList ("-NoProfile -NoLogo -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
  Exit	
}
#hide powershell console
Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

function Show-ModernFilePicker {
  param(
    [ValidateSet('Folder', 'File')]
    $Mode,
    [string]$fileType

  )

  if ($Mode -eq 'Folder') {
    $Title = 'Select Folder'
    $modeOption = $false
    $Filter = "Folders|`n"
  }
  else {
    $Title = 'Select File'
    $modeOption = $true
    if ($fileType) {
      $Filter = "$fileType Files (*.$fileType) | *.$fileType|All files (*.*)|*.*"
    }
    else {
      $Filter = 'All Files (*.*)|*.*'
    }
  }
  #modern file dialog
  #modified code from: https://gist.github.com/IMJLA/1d570aa2bb5c30215c222e7a5e5078fd
  $AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
  $Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.AddExtension = $modeOption
  $OpenFileDialog.CheckFileExists = $modeOption
  $OpenFileDialog.DereferenceLinks = $true
  $OpenFileDialog.Filter = $Filter
  $OpenFileDialog.Multiselect = $false
  $OpenFileDialog.Title = $Title
  $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

  $OpenFileDialogType = $OpenFileDialog.GetType()
  $FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
  $IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null)
  $null = $OpenFileDialogType.GetMethod('OnBeforeVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $IFileDialog)
  if ($Mode -eq 'Folder') {
    [uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
    $FolderOptions = $OpenFileDialogType.GetMethod('get_Options', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null) -bor $PickFoldersOption
    $null = $FileDialogInterfaceType.GetMethod('SetOptions', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $FolderOptions)
  }
  
  

  $VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName, 'System.Windows.Forms.FileDialog+VistaDialogEvents', $false, 0, $null, $OpenFileDialog, $null, $null).Unwrap()
  [uint32]$AdviceCookie = 0
  $AdvisoryParameters = @($VistaDialogEvent, $AdviceCookie)
  $AdviseResult = $FileDialogInterfaceType.GetMethod('Advise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdvisoryParameters)
  $AdviceCookie = $AdvisoryParameters[1]
  $Result = $FileDialogInterfaceType.GetMethod('Show', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, [System.IntPtr]::Zero)
  $null = $FileDialogInterfaceType.GetMethod('Unadvise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdviceCookie)
  if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileDialogInterfaceType.GetMethod('GetResult', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $null)
  }

  return $OpenFileDialog.FileName
}


function Get-PowerPlanSettings {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Guid
  )

  if (-not ($Guid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')) {
    Update-Log -msg 'Invalid GUID format.' -type error
    return
  }

  #get all power plan settings with undocumented /qh arg
  $result = powercfg.exe /qh $Guid
  if (-not $result) {
    Update-Log -msg "Failed to retrieve power plan settings for GUID: $Guid" -type error
    return
  }

  #get setting description
  $wmiSettings = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerSetting
  if (-not $wmiSettings) {
    Update-Log -msg 'Failed to retrieve WMI power settings. Descriptions will be unavailable.' -type error
  }
  #get power plan description
  $planDesc = (Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "InstanceID LIKE '%$Guid%'" -ErrorAction SilentlyContinue).Description

  $powerPlan = [PSCustomObject]@{
    PowerSchemeGUID = $Guid
    PowerSchemeName = $null
    Description     = $null
    Subgroups       = @()
  }

  if ($planDesc) {
    $powerPlan.Description = $planDesc
  }

  $currentSubgroup = $null
  $currentSetting = $null
  $readSettings = $false
  $currentIndex = $null

  foreach ($line in $result) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    #get power scheme name
    if ($line -like '*Power Scheme GUID*') {
      $splitLine = $line -split ':'
      $name = ($splitLine[1].Trim() -replace '.*\((.*)\).*', '$1').Trim()
      $powerPlan.PowerSchemeName = $name
      continue
    }

    #get subgroups
    if ($line -like '*Subgroup GUID*') {
      $splitLine = $line -split ':'
      $subGuid = ($splitLine[1].Trim() -split '\s+')[0]
      $subName = ($splitLine[1].Trim() -replace '.*\((.*)\).*', '$1').Trim()
            
      $currentSubgroup = [PSCustomObject]@{
        SubgroupGUID = $subGuid
        SubgroupName = $subName
        Settings     = @()
      }
      $powerPlan.Subgroups += $currentSubgroup
      $readSettings = $false
      continue
    }

    #get power setting info
    if ($line -like '*Power Setting GUID*') {
      $readSettings = $true
      $splitLine = $line -split ':'
      $settingGuid = ($splitLine[1].Trim() -split '\s+')[0]
      $settingName = ($splitLine[1].Trim() -replace '.*\((.*)\).*', '$1').Trim()

      $desc = ($wmiSettings | Where-Object { $_.InstanceID -like "*{$settingGuid}" } | Select-Object -First 1).Description 

      $currentSetting = [PSCustomObject]@{
        PowerSettingGUID = $settingGuid
        PowerSettingName = $settingName
        Description      = $desc
        PossibleValues   = @()  # For named options (Index and FriendlyName)
        MinPossible      = $null  # For numeric ranges
        MaxPossible      = $null  # For numeric ranges
        Increment        = $null  # For numeric ranges
        Units            = $null  # For numeric ranges
        CurrentACValue   = $null
        CurrentDCValue   = $null
        IsNumericRange   = $false  
        SubgroupGUID     = $currentSubgroup.SubgroupGUID 
      }
      if ($currentSubgroup) {
        $currentSubgroup.Settings += $currentSetting
      }
      continue
    }

    #for named settings
    if ($readSettings -and $line -like '*Possible Setting Index*') {
      $splitLine = $line -split ':'
      $currentIndex = [Convert]::ToInt32($splitLine[1].Trim(), 16) 
      continue
    }

    #friendly Name for named settings
    if ($readSettings -and $line -like '*Possible Setting Friendly Name*') {
      $splitLine = $line -split ':'
      $friendlyName = $splitLine[1].Trim()
      if ($null -ne $currentIndex) {
        $currentSetting.PossibleValues += [PSCustomObject]@{
          Index        = $currentIndex
          FriendlyName = $friendlyName
        }
        $currentIndex = $null 
      }
      continue
    }

    #min setting for numerics
    if ($readSettings -and $line -like '*Minimum Possible Setting*') {
      $splitLine = $line -split ':'
      $currentSetting.MinPossible = [Convert]::TouInt32($splitLine[1].Trim(), 16)
      $currentSetting.IsNumericRange = $true
      continue
    }

    #max setting for numerics
    if ($readSettings -and $line -like '*Maximum Possible Setting*') {
      $splitLine = $line -split ':'
      $currentSetting.MaxPossible = [Convert]::TouInt32($splitLine[1].Trim(), 16)
      continue
    }

    #get numeric increment 
    if ($readSettings -and $line -like '*Possible Settings increment*') {
      $splitLine = $line -split ':'
      $currentSetting.Increment = [Convert]::TouInt32($splitLine[1].Trim(), 16)
      continue
    }

    #get numeric unit
    if ($readSettings -and $line -like '*Possible Settings units*') {
      $splitLine = $line -split ':'
      $currentSetting.Units = $splitLine[1].Trim()
      continue
    }

    #current ac setting value
    if ($readSettings -and $line -like '*Current AC Power Setting Index*') {
      $splitLine = $line -split ':'
      $hexValue = $splitLine[1].Trim()
      $currentSetting.CurrentACValue = [Convert]::TouInt32($hexValue, 16)
      continue
    }

    #current dc setting value
    if ($readSettings -and $line -like '*Current DC Power Setting Index*') {
      $splitLine = $line -split ':'
      $hexValue = $splitLine[1].Trim()
      $currentSetting.CurrentDCValue = [Convert]::TouInt32($hexValue, 16)
      continue
    }
  }

  return $powerPlan
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Power Plan Settings Editor'
$form.Size = New-Object System.Drawing.Size(1300, 690)
$form.Font = New-Object System.Drawing.Font('segoe ui', 8)
$form.StartPosition = 'CenterScreen'
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAG+SURBVDiNdZMxSBthGIaf778jVzyhFmLJILQ6FEooLsVNIbVCO7d0sEuLgji0BYmjdC4ougQdzODSxT0Fa0uFLp1axGQoCkELKhYsRAl3zeV3+M8kl0s+OHjvfd/v5fv/+05oK72TeAAyjWYc4a4hKSNsU9frkvH3Wv3SaCzg0Osso5kBVHtwWAGwRtKbkzR+I0AXcHCdT0AmYh/6AM6AwdV9KL+/Vr6S9J5KGt8GwHVWYs1uGgbeNd+Lz1rVR/x1lsB7I+GZf8XGvrcKqdcG//sGu0/ix7H0sAKZjjXbfdD/wmAdwMF8p/uwqMmUAh7HpNQrsFyDzzbB+wP2LRA76hMmRO84FaC3ySoYKcKNwaj5cg9+jkK92spWbEBHx78JR4sGJ1JwZwGCCyhNtjcDaAUcRajaORznzRNcGG7/LVR/d7qHQwV87qQAcHsSTjbg9GNnXdhS1PU6ZsOi1XMfrB44yHaLD0DnVbjbazG5/zmUXkJQ6RaQkzG/ZL5/0psDvjQ1BZUfcLnbbfRtXC9rYFi6SMKsJ7OIbaFrXcYmh+tl5SH/IwGNoO+JNDWZQpiA8HeGMsIW6LyM+aVW/xUNWYWRwFJ7iQAAAABJRU5ErkJggg=='
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$form.icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($memoryStream).GetHIcon()))
$memoryStream.Dispose()
$form.ShowInTaskbar = $true

$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Location = New-Object System.Drawing.Point(10, 10)
$treeView.Size = New-Object System.Drawing.Size(250, 500)
$treeView.BackColor = [System.Drawing.Color]::FromArgb(60, 80, 121) #rgb(60, 80, 121)
$treeView.ForeColor = 'White'
$treeView.HideSelection = $false
$treeView.Cursor = 'Hand'
$treeView.Font = New-Object System.Drawing.Font('segoe ui', 9)
$treeView.ItemHeight = 20
$treeView.LineColor = [System.Drawing.Color]::FromArgb(21, 24, 34) #rgb(21, 24, 34)
$treeView.DrawMode = [System.Windows.Forms.TreeViewDrawMode]::OwnerDrawText
#custom select color
$treeView.Add_DrawNode({
    param($sender, $e)
    $node = $e.Node

    if ($node.IsSelected) {
      $selectionColor = [System.Drawing.Color]::FromArgb(89, 100, 156) #rgb(89, 100, 156)
      $brush = New-Object System.Drawing.SolidBrush($selectionColor)
      $e.Graphics.FillRectangle($brush, $e.Bounds)
      $brush.Dispose()
    }
    else {
      $e.Graphics.FillRectangle([System.Drawing.Brushes]::Transparent, $e.Bounds)
    }

    $textBrush = New-Object System.Drawing.SolidBrush($treeView.ForeColor)
    $e.Graphics.DrawString($node.Text, $treeView.Font, $textBrush, $e.Bounds.X, $e.Bounds.Y)
    $textBrush.Dispose()
  })

$form.Controls.Add($treeView)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(260, 10)
$dataGridView.Size = New-Object System.Drawing.Size(1000, 500)
$dataGridView.RowHeadersVisible = $false

$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(219, 219, 219)
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.GridColor = [System.Drawing.Color]::FromArgb(100, 100, 100) #rgb(180, 180, 180)
$dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dataGridView.AllowUserToResizeRows = $true
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(160, 197, 231) #rgb(160, 197, 231)
$dataGridView.DefaultCellStyle.SelectionForeColor = 'Black'
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.ReadOnly = $true
$dataGridView.Cursor = 'Hand'

$editColumn = New-Object System.Windows.Forms.DataGridViewImageColumn
$editColumn.Name = 'Edit'
$editColumn.HeaderText = ''
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAXAAAAFwBhyfMcAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAACXSURBVDiNzdI9CsJQEEXhzyI7CP5sxYXYuA4t3JO9XbS21ARBXILgAoRYJA9CiOElNh6YZuDcO8UwjBkyPLEZ6EpwQdmYbayc1gFrvBsBrxh5iiv2HSFZrBwaQ8iq3i+GyO2QpE+eo+iQy3qfjmkucRt79p/LCfIvcqF6316WY5sDu19kOLTkPObswAQP3HHCEWfVr0fxAbSiWClHz+v4AAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$editColumn.Image = [System.Drawing.Image]::FromStream($memoryStream)
$editColumn.ImageLayout = [System.Windows.Forms.DataGridViewImageCellLayout]::Normal
$memoryStream.Dispose()
$dataGridView.Columns.Add($editColumn)

for ($count = 1; $count -lt 7; $count++) {
  $textCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $dataGridView.Columns.Add($textCol)
}
$dataGridView.Columns[1].Name = 'Name'
$dataGridView.Columns[1].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(193, 233, 233) #rgb(193, 233, 233)
$dataGridView.Columns[2].Name = 'Description'
$dataGridView.Columns[2].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(205, 255, 200) #rgb(205, 255, 200)
$dataGridView.Columns[3].Name = 'Value AC (Plugged In)'
$dataGridView.Columns[4].Name = 'Value DC (On Battery)'
$dataGridView.Columns[5].Name = 'Min Value'
$dataGridView.Columns[6].Name = 'Max Value'

$dataGridView.Columns[0].Width = 40
$dataGridView.Columns[1].Width = 200    
$dataGridView.Columns[2].Width = 400    
$dataGridView.Columns[3].Width = 130    
$dataGridView.Columns[4].Width = 130    
$dataGridView.Columns[5].Width = 100    
$dataGridView.Columns[6].Width = 100    

$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('segoe ui', 8, [System.Drawing.FontStyle]::Bold)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None 
$dataGridView.AllowUserToAddRows = $false
$form.Controls.Add($dataGridView)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 520)
$comboBox.Size = New-Object System.Drawing.Size(250, 20)
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboBox)
function Pop-ComboBox {
  $powerPlansRaw = powercfg /l
  $comboBox.Items.Clear()
  if (-not $powerPlansRaw) {
    [System.Windows.Forms.MessageBox]::Show('Failed to retrieve power plans. Ensure the script is run as Administrator.', 'Error')
    return
  }

  $powerPlans = $powerPlansRaw | Where-Object { $_ -match '(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})\s+\((.*?)\)' } | ForEach-Object {
    $guid = $matches[1]
    $name = $matches[2]
    [PSCustomObject]@{ GUID = $guid; Name = "$name" }
  }

  if ($powerPlans.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show('No power plans found. Ensure the script is run as Administrator.', 'Error')
    return
  }

  foreach ($plan in $powerPlans) {
    if ($plan.guid -notin $comboBox.Items.GUID) {
      $comboBox.Items.Add($plan)
    }
    
  }
  $comboBox.SelectedIndex = 0
  $comboBox.DisplayMember = 'Name'
}
Pop-ComboBox

$restoreButton = New-Object System.Windows.Forms.Button
$restoreButton.Location = New-Object System.Drawing.Point(600, 520)
$restoreButton.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAAqBJREFUOI1tk02I1WUUxs/9e1VGaWxGwlVqCBLdZCQHdRNYMsgkEehMhFoEk4qKCBIEru4yBT/SgsjxixZFEe0c9FquLATFr8aFkowU0QWdmRy53P95nue8LuY/w23qbN538TznHH7nnJLNiP5aWuHOj4hYL8RSQkaPEVKXAmnwl12dv7XqS1Of3vNp7rzZcZRIO+nKhLDCbCpeQoKnL9F4tH+4WvHpBJPmNEREN13PCTFK53d0uyWhzDy9SrCfnjoFGRE/l+xJ73C14mUzs3mz4xiR3pDrPSFennjy9MiVgRcmWltd9enoJ2LjBJnep8eb8rbDZra31F9LKxy6SVcW5GtDffNvrLucyvx9/APlsUFUhjyGbjQWnbVqKZYfePi1oG30UGLqKgMcoEcmhIlWNzMr3xubnzPtILSmYND3iv3Z17bj2juPG7E3pdhIqIOIgQwePVPAkGvw9ZNjVxtMdwitnAaIMHr0jrd3fjhy7KVxIX6kh4WzJxNicQvtXnqsBvQiPea2mItEWmdmBo9bggyelmSEUiE4Q8S7hK7MqDw9RlBmZhbA1Fgtg8cfhcB/3b3we3ic/T8zIVMzLk92kLo4qXmYiaoVgr41xx+332wsOk3EhZlmIi7WZzXOLdhyu4PQJiFMYC0LpEFCosfC5kTzM6uWou2v+tti7KTHt4K+Ya7t9ft/b7SvuiHXCSGep0vI02DJzGzlofrnAe4hwog4R3Lfg4PL/mldpAVbbncU5q10mRjHo7Z2X8nMrFIdnpOs/Tw91k+2rVEgfhDiTiAMTXaRsbmobERcSmPlt+x6N6aPqVIdntN82naYiF2EZv0HIMLokhhfxGj5Y7vejX9d41Qs3nO/kjsHwtkDT0sLgCPK4yKgU/bT2rut+medyZK38MnmgAAAAABJRU5ErkJggg=='
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$restoreButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$restoreButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$restoreButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$restoreButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$restoreButton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$restoreButton.Text = ' Restore All'
$restoreButton.Cursor = 'Hand'
$form.Controls.Add($restoreButton)

$tooltip1 = New-Object System.Windows.Forms.ToolTip
$tooltip1.SetToolTip($restoreButton, 'Restores Installed Power Plans to Default')

$restoreButton.Add_Click({
    Update-Log -msg 'Restoring Default Plans...' -type output
    powercfg /restoredefaultschemes
    Pop-ComboBox
    Update-Display
  })

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(710, 520)
$refreshButton.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAAeJJREFUOI19kjFoVEEQhr/Z3btc1DMkKkTQFIIGtDIRLLQSC7XQxoiIWChCtFBsBBEkEMEYKyFobCUW2gUFLQRBjdpZnRATFCWFFl7QRO923x1r8bJ7z8uZhYV5s7Pf/O/fEVZYXa98Cdi+7EC4Vt4rwwBmJYC2FJsuAjztsIyWl1L/ANY/8xsLsHrugMwCGNeS+2X2kNjwoULQ88R3Fqq8UBU6ogLXcg9umvRnlwFkkdvG0ltMmA45Y/+zq9zY8MiviYAt932PSTihHbXSB/5kPEA5HivLOe2oZ1SsKzoOR0DOcVxbtLaY3s10RwUJV9rbOPrplIwrxyVtU6hOleyJAOPYZlxqmqqyPwCmz8iD0jFxADOnGctZ3oY6bdNGCkA5uoI8k3CeIR+9aTyheJ0wHuscPxoKLN8zJu3uW8VFgB1DPt834k9mPHke6rRlBpbmwFimvDCY6Xerf9hXxXOQhJ3ABMD7Ct92aWqA+DoPIyC/yGQ9zzzQGZoBd1LpzAVqPxSMxQAT767L1/gLU6OyoBKuthocE2cO2its1Y6PpsaFkItmvbkpd3OWe8Hl6HZmnNVvfrUl7Hs9IvPxqUMwMODzZUu3lyb3fSN8OSafm04bgJ9ruawTjjQXeFhozmXXX6Fwt+T7PRKyAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$refreshButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$refreshButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$refreshButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$refreshButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$refreshButton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$refreshButton.Text = ' Refresh'
$refreshButton.Cursor = 'Hand'
$form.Controls.Add($refreshButton)

$tooltip2 = New-Object System.Windows.Forms.ToolTip
$tooltip2.SetToolTip($refreshButton, 'Refresh the Display to Update Any Chnages')

$removePlanButton = New-Object System.Windows.Forms.Button
$removePlanButton.Location = New-Object System.Drawing.Point(270, 520)
$removePlanButton.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAASlJREFUOI21kb8uRFEQxn+zlz0nKzT+JBLNbiUSr6DzCCLxABKJUiUS2/AIZIst6XS04gmoNCqFKCiQ2HvWxr2juMc6rnsRf75kciYzZ775ZgZ+CfkQaengZI2mcUQmBuvAZJZEQ2yeNOU5/D6Qrx8bp57GnF8uy34Yn1vTpeiWBnBRqKByrFvGsWIdkcmsZx0EKqrGkVhHYrrsHOzJRl5B3VMmQFIwbi/wG4U7qB3ptulS6XcNdmBjqD6BjUkP27JeuAOFVZRh4EZhAuEhzZqMqHCFMgU8An2CSk5m7N9dr+8MOAUQoe1znbAgT9Dha5QT6JuCUoh8QiDfUKD6nyP8hODdGQXuvDMDpCijkp0xVWXa5+5LCVRpAfMoCz40G6QXgWt5PfFf4QVER2H583WohwAAAABJRU5ErkJggg=='
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$removePlanButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$removeplanbutton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$removeplanbutton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$removeplanbutton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$removeplanbutton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$removePlanButton.Text = ' Delete Plan'
$removePlanButton.Cursor = 'Hand'
$form.Controls.Add($removePlanButton)

$tooltip3 = New-Object System.Windows.Forms.ToolTip
$tooltip3.SetToolTip($removePlanButton, 'Delete Currently Selected Plan')

$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(380, 520)
$importButton.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAARxJREFUOI3NkL1KA1EQhb/Z3L1LtBGsRFIIltrZCD6ADyCWFoKFiiDGJuQBJEUQ/GsMKSx9iFTaWalgk04kVqKFyd6rMBZZJZuQRbTQA8PAnOG79wz8a5mGVkxDK5k7mWbMFMD7TwHWd3v8W0CWzPSBRq9jzLZW5GrAzHh64kznRp+5MS7PovGcA9GnWajpLtAm+UHhVNeBkfs1qX7BPZcuz7KJ3rCq2F565GgrHAFPyWhJhK1UPIcVwQY5D2Ff1uamnFhHMfSMJ1Vubshx707oIefB2A4ggxnvdmR/Zk9fAgiuS1Lr90MHKJisS9+WpD7Ms67bTS7mMQhgvqzbwMNwXEqT6kGhJaCyUORQhFUg/01AR5X6RTV92L/RB6+ZVgEny4OPAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$importButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$importButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$importButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$importButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$importButton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$importButton.Text = ' Import'
$importButton.Cursor = 'Hand'
$form.Controls.Add($importButton)

$tooltip4 = New-Object System.Windows.Forms.ToolTip
$tooltip4.SetToolTip($importButton, 'Import Power Plan from POW File')

$importButton.Add_Click({
    $plan = Show-ModernFilePicker -Mode File -fileType 'pow'
    if ($plan) {
      Update-Log -msg 'Importing Plan...' -type output
      $guid = New-Guid
      powercfg /import $plan $guid
      Pop-ComboBox
      Update-Display
    }

  })

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(490, 520)
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAAcFJREFUOI19kz1oFEEUx3/vZuaWjeIdWFiLZQQLRRARBNHSqHBdYhFEBLWz1tgKVloYCRZiquuilYIERIQQK6OFoFXSKkETZXdmnsXs3W3uznvNe7Nv5//xZkbOPdaDmjFnLJm10LDQyFJ21dpmIJat5dPygqGwhWfeCA/Q9MEACIiCr62NwtV3uvb8jHzdAxBK9lX1Z1Heqw5zADAH5FFpjiiIPjGosLp6U26N231pRWeA3MbRnvVFkiljiVNEz20R2ruB76MKAlAC46UDsHJFur368is9emydLwsLEgcWGAAcX1SXB/b3NrRaNaQWaOT+p5M0O290tntetpOFGrv9wZPomDcWjIM/OykbC+Y3RNf/da3zVi8OFFQRygnDqIdCQ8B6nxRIpUI3ubF7mDutErQE3yRdCJ/sBM+SQDMLzHZnZNvGAnCDGXx8KiXw83/EF5b03onNoSFOOAAATj3UjjW02znLL6/Jxutar29hEkgseBQch3795QOwUe/1LTTg7JHrumgcWAdSZeMglBwAKMawWF+yY1I9LTDdaxiFoCRp1dGpUowAmMizUJIrTCEgArHGpNC759/W7+59iQD/AI2kp+UrwN4nAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$exportButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$exportButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$exportButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$exportButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$exportButton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$exportButton.Text = ' Export'
$exportButton.Cursor = 'Hand'
$form.Controls.Add($exportButton)

$tooltip5 = New-Object System.Windows.Forms.ToolTip
$tooltip5.SetToolTip($exportButton, 'Export Current Plan to POW File')

$exportButton.Add_Click({
    $dest = Show-ModernFilePicker -Mode Folder
    if ($dest) {
      $randomNum = Get-Random -Minimum 100 -Maximum 1000
      Update-Log -msg "Exporting Plan to $dest\$($comboBox.SelectedItem.Name)-BACKUP$randomNum.pow..." -type output
      powercfg /export "$dest\$($comboBox.SelectedItem.Name)-BACKUP$randomNum.pow" $comboBox.SelectedItem.GUID
    }

  })

$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Location = New-Object System.Drawing.Point(10, 550)
$groupBox.Size = New-Object System.Drawing.Size(630, 120) 
$groupBox.Text = 'Power Plan Details'
$form.Controls.Add($groupBox)

$labelName = New-Object System.Windows.Forms.Label
$labelName.Text = 'Name'
$labelName.Location = New-Object System.Drawing.Point(10, 20)
$labelName.Size = New-Object System.Drawing.Size(100, 20)
$groupBox.Controls.Add($labelName)

$labelDesc = New-Object System.Windows.Forms.Label
$labelDesc.Text = 'Description'
$labelDesc.Location = New-Object System.Drawing.Point(240, 20)
$labelDesc.Size = New-Object System.Drawing.Size(100, 20)
$groupBox.Controls.Add($labelDesc)

$textBoxPlanName = New-Object System.Windows.Forms.TextBox
$textBoxPlanName.Location = New-Object System.Drawing.Point(10, 40)
$textBoxPlanName.Size = New-Object System.Drawing.Size(230, 20)
$textBoxPlanName.BackColor = [System.Drawing.Color]::FromArgb(193, 233, 233) 
$textBoxPlanName.BorderStyle = [system.windows.forms.BorderStyle]::Fixed3D
$groupBox.Controls.Add($textBoxPlanName)

$textBoxDescription = New-Object System.Windows.Forms.TextBox
$textBoxDescription.Location = New-Object System.Drawing.Point(240, 40)
$textBoxDescription.Size = New-Object System.Drawing.Size(300, 20)
$textBoxDescription.BackColor = [System.Drawing.Color]::FromArgb(205, 255, 200)
$textBoxDescription.BorderStyle = [system.windows.forms.BorderStyle]::Fixed3D
$groupBox.Controls.Add($textBoxDescription)

$checkBoxActive = New-Object System.Windows.Forms.CheckBox
$checkBoxActive.Location = New-Object System.Drawing.Point(560, 40)
$checkBoxActive.Size = New-Object System.Drawing.Size(60, 20)
$checkBoxActive.Text = 'Active'
$groupBox.Controls.Add($checkBoxActive)
$checkBoxActive.Add_Click({
    if (!($checkBoxActive.Checked)) {
      Update-Log -msg 'Setting Different Plan Active...' -type output
      $nextGUID = ($comboBox.Items | Where-Object { $_.GUID -ne $comboBox.SelectedItem.GUID } | Select-Object -First 1).GUID
      powercfg /setactive $nextGUID
    }
    else {
      Update-Log -msg 'Setting Current Plan Active...' -type output
      powercfg /setactive $comboBox.SelectedItem.GUID
    }
  })

$buttonUpdate = New-Object System.Windows.Forms.Button
$buttonUpdate.Location = New-Object System.Drawing.Point(520, 70)
$buttonUpdate.Size = New-Object System.Drawing.Size(100, 30)
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAs/AAALPwFJxTL7AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAARxJREFUOI2l078uhEEUBfDffLZkOx6AjUrhCTSyJBsKCQolEQ2JN1HoFDqNf42QrEqvVmCjUy3NZgvFJp/CLMN+uyRON2fuuefMzB3+idDDXBlVsoMFQSWyj3KXOvbVNPs3qFuWOZQb6WPYEqyrOuttULcsOC5M9R25YKXbJETxmExjgHNvko6KmmYWc2z/Iu4IDvAa1+V4T7JILA4Us6ZqC3cJv5g2GE82GtjEy6d4zom6PcwkdRUoFTi+GXak5daQCVXn6vYEu0XxugmeEm5K24Wyh1/EjTTBBaaTzVltN649Y6nIOWriM35MXwPlPsU/8eMZa5pyG8j/IM4F692Rzj7peadyq2gNdE6m8OsIKdLPxGSsuu/3mf6NdwQ5Tr+kaxstAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$buttonUpdate.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$buttonUpdate.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$buttonUpdate.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$buttonUpdate.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
$buttonUpdate.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$buttonUpdate.Text = ' Update'
$buttonUpdate.Cursor = 'Hand'
$groupBox.Controls.Add($buttonUpdate)

$tooltip6 = New-Object System.Windows.Forms.ToolTip
$tooltip6.SetToolTip($buttonUpdate, 'Save Changes to Power Plan Name and Description')

#populate tree and data grid
function Update-Display {
  Update-Log -msg 'Updating Display...' -type output
  $selectedNode = $treeView.SelectedNode.Tag
  $treeView.Nodes.Clear()
  $dataGridView.Rows.Clear()
  $selectedPlan = $comboBox.SelectedItem
  if (-not $selectedPlan) {
    [System.Windows.Forms.MessageBox]::Show('No power plan selected.', 'Error')
    return
  }

  $Global:settings = Get-PowerPlanSettings -Guid $selectedPlan.GUID
  if (-not $Global:settings) {
    [System.Windows.Forms.MessageBox]::Show("Failed to retrieve settings for GUID: $($selectedPlan.GUID). Ensure the script is run as Administrator.", 'Error')
    return
  }

  if (-not $Global:settings.Subgroups) {
    [System.Windows.Forms.MessageBox]::Show('No subgroups found for the selected power plan.', 'Warning')
    return
  }

  $textBoxPlanName.Text = $selectedPlan.Name
  $textBoxDescription.Text = $Global:settings.Description 
  $activeScheme = (powercfg /getactivescheme | Select-String -Pattern '\w{8}-\w{4}-\w{4}-\w{4}-\w{12}').Matches.Value
  $checkBoxActive.Checked = ($activeScheme -eq $selectedPlan.GUID)

  foreach ($subgroup in $Global:settings.Subgroups) {
    $node = $treeView.Nodes.Add($subgroup.SubgroupName)
    $node.Tag = $subgroup
    #select previously selected node
    if ($selectedNode -and $subgroup.SubgroupGUID -eq $selectedNode.SubgroupGUID) {
      $treeView.SelectedNode = $node
    }
  }

  if (!($treeView.SelectedNode)) {
    #no previously selected node so default to first subgroup
    if ($Global:settings.Subgroups.Count -gt 0) {
      $firstSubgroup = $Global:settings.Subgroups[0]
      if (-not $firstSubgroup.Settings) {
        [System.Windows.Forms.MessageBox]::Show("No settings found in the first subgroup: $($firstSubgroup.SubgroupName).", 'Warning')
        return
      }

      foreach ($setting in $firstSubgroup.Settings) {
        $minValue = 'N/A'
        $maxValue = 'N/A'
        $acValue = $setting.CurrentACValue
        $dcValue = $setting.CurrentDCValue

        if (!($setting.PossibleValues) -and $setting.IsNumericRange -eq $false) {
          $setting.IsNumericRange = $true
          $setting.MinPossible = 0
          $setting.MaxPossible = 255
        }

        if ($setting.IsNumericRange) {
          $minValue = "$($setting.MinPossible) $($setting.Units)"
          $maxValue = "$($setting.MaxPossible) $($setting.Units)"
          $acValue = "$($setting.CurrentACValue) $($setting.Units)"
          $dcValue = "$($setting.CurrentDCValue) $($setting.Units)"
        }
        elseif ($setting.PossibleValues.Count -gt 0) {
          $indices = $setting.PossibleValues | ForEach-Object { [int]$_.Index }
          $minIndex = $indices | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
          $maxIndex = $indices | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

          $minValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $minIndex } | Select-Object -First 1).FriendlyName
          $maxValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $maxIndex } | Select-Object -First 1).FriendlyName

          $acFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentACValue } | Select-Object -First 1).FriendlyName
          $dcFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentDCValue } | Select-Object -First 1).FriendlyName
          $acValue = if ($acFriendly) { $acFriendly } else { $setting.CurrentACValue }
          $dcValue = if ($dcFriendly) { $dcFriendly } else { $setting.CurrentDCValue }
        }

        $rowIndex = $dataGridView.Rows.Add(
          $editColumn.Image, 
          $setting.PowerSettingName,
          $setting.Description,
          $acValue,
          $dcValue,
          $minValue,
          $maxValue
        )
        #store the object in the rows tag for later use
        $dataGridView.Rows[$rowIndex].Tag = $setting
      }
    }
  }
  $dataGridView.ClearSelection()
}


$buttonUpdate.Add_Click({
    Update-Log -msg 'Updating Power Plan Name and Description...' -type output
    $newName = $textBoxPlanName.Text
    $newDescription = $textBoxDescription.Text
    $selectedGuid = $comboBox.SelectedItem.GUID
    powercfg /changename $selectedGuid $newName $newDescription
    Pop-ComboBox  
    Update-Display
  })

$removePlanButton.Add_Click({
    Update-Log -msg 'Removing Current Plan...' -type output
    powercfg /delete $comboBox.SelectedItem.GUID
    Pop-ComboBox
    Update-Display
  })

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$editSetting = New-Object System.Windows.Forms.ToolStripMenuItem
$editSetting.text = 'Edit'
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAB2AAAAdgFOeyYIAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAAERJREFUOI1jYBjsIAOKyQbHoRgnYKLE9MFhACMSO4OBgSEFTV4DSt9AE5/DwMAwA90F/0mwmGi1IyEWWAjIL6TUAtoDAPw4CZ9G70B7AAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$editSetting.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$contextMenu.Items.Add($editSetting)
$lookupSetting = New-Object System.Windows.Forms.ToolStripMenuItem
$lookupSetting.text = 'Look Up'
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAQlAAAEJQGmsd4JAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAARZJREFUOI2d078rhXEUx/GXmxTDNVyKUpKU/0MyWRXFILHKYmK2mEwGGY0mg8nCoBgMBiQGv7pXMRrkGr7Hryf3eeTU06lzvp/39/z4Pvy0PqzhAi+oYRfjKCmw6RDVcYkd7H+L7aHSSDyGN1xjKJOrYD0gB2jJistR6gN6cipcDch8NjETibkcMbTiHufZxGYAOgoAsBFnuz4CJanHVzz+AXAX/vOyEqpoRvcfAL3hq9+Dk1HWQoG4HFWeZhNtuMETBhuIm3zNava3AyPSHGqYkFr6sH5sh/hKzoscjSrqeMaxtLJ6fNXwK40ApI0sSS/uFmfYwrA0g8OALOdB8qwdRwFZ/C+kU9rEG6YKf9FfrCa1dIKBd5MKRCD5YQKsAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$lookupSetting.Image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$contextMenu.Items.Add($lookupSetting)
$dataGridView.ContextMenuStrip = $contextMenu 


function Edit-Setting {
  param(
    $row
  )
   
  $setting = $row

  $editForm = New-Object System.Windows.Forms.Form
  $editForm.Text = 'Edit Setting'
  $editForm.Size = New-Object System.Drawing.Size(440, 430)  
  $editForm.StartPosition = 'CenterParent'
  $editForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $editForm.Font = New-Object System.Drawing.Font('segoe ui', 8)
  $editForm.MaximizeBox = $false
  $editForm.MinimizeBox = $false

  $labelName = New-Object System.Windows.Forms.Label
  $labelName.Text = $setting.PowerSettingName
  $labelName.Location = New-Object System.Drawing.Point(10, 10)
  $labelName.Size = New-Object System.Drawing.Size(425, 40)
  $labelName.Font = New-Object System.Drawing.Font('segoe ui', 10, [System.Drawing.FontStyle]::Bold)
  $editForm.Controls.Add($labelName)

  $textDesc = New-Object System.Windows.Forms.TextBox
  $textDesc.Text = $setting.Description 
  $textDesc.Location = New-Object System.Drawing.Point(10, 60)
  $textDesc.Size = New-Object System.Drawing.Size(380, 60)
  $textDesc.BackColor = [System.Drawing.Color]::FromArgb(205, 255, 200)
  $textDesc.BorderStyle = [system.windows.forms.BorderStyle]::Fixed3D
  $textDesc.Multiline = $true
  $textDesc.TabStop = $false
  $textDesc.ReadOnly = $true
  $textDesc.WordWrap = $true
  $editForm.Controls.Add($textDesc)

  $minValue = 'N/A'
  $maxValue = 'N/A'
  #fix a few settings that are numeric values but arent reported as such in powercfg
  if (!($setting.PossibleValues) -and $setting.IsNumericRange -eq $false) {
    $setting.IsNumericRange = $true
    $setting.MinPossible = 0
    $setting.MaxPossible = 255
  }
  if ($setting.IsNumericRange) {
    $minValue = "$($setting.MinPossible) $($setting.Units)"
    $maxValue = "$($setting.MaxPossible) $($setting.Units)"
  }
  elseif ($setting.PossibleValues.Count -gt 0) {
    $indices = $setting.PossibleValues | ForEach-Object { [int]$_.Index }
    $minIndex = $indices | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
    $maxIndex = $indices | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
    $minValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $minIndex } | Select-Object -First 1).FriendlyName
    $maxValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $maxIndex } | Select-Object -First 1).FriendlyName
  }

  $labelMinMax = New-Object System.Windows.Forms.Label
  $labelMinMax.Text = "Min: $minValue  Max: $maxValue"
  $labelMinMax.Location = New-Object System.Drawing.Point(10, 130)
  $labelMinMax.Size = New-Object System.Drawing.Size(420, 20)
  $labelMinMax.Font = New-Object System.Drawing.Font('segoe ui', 8, [System.Drawing.FontStyle]::Bold)
  $editForm.Controls.Add($labelMinMax)

  $labelAC = New-Object System.Windows.Forms.Label
  $labelAC.Text = 'Value AC (Plugged In):'
  $labelAC.Location = New-Object System.Drawing.Point(10, 160)
  $labelAC.Size = New-Object System.Drawing.Size(150, 20)
  $editForm.Controls.Add($labelAC)

  $labelDC = New-Object System.Windows.Forms.Label
  $labelDC.Text = 'Value DC (On Battery):'
  $labelDC.Location = New-Object System.Drawing.Point(10, 200)
  $labelDC.Size = New-Object System.Drawing.Size(150, 20)
  $editForm.Controls.Add($labelDC)

  if ($setting.IsNumericRange) {
    $numericAC = New-Object System.Windows.Forms.NumericUpDown
    $numericAC.Location = New-Object System.Drawing.Point(160, 160)
    $numericAC.Size = New-Object System.Drawing.Size(150, 20)
    $numericAC.Minimum = $setting.MinPossible
    $numericAC.Maximum = $setting.MaxPossible
    $numericAC.Increment = $setting.Increment
    $numericAC.Value = $setting.CurrentACValue
    $editForm.Controls.Add($numericAC)

    $numericDC = New-Object System.Windows.Forms.NumericUpDown
    $numericDC.Location = New-Object System.Drawing.Point(160, 200)
    $numericDC.Size = New-Object System.Drawing.Size(150, 20)
    $numericDC.Minimum = $setting.MinPossible
    $numericDC.Maximum = $setting.MaxPossible
    $numericDC.Increment = $setting.Increment
    $numericDC.Value = $setting.CurrentDCValue
    $editForm.Controls.Add($numericDC)
  }
  else {
    $comboAC = New-Object System.Windows.Forms.ComboBox
    $comboAC.Location = New-Object System.Drawing.Point(160, 160)
    $comboAC.Size = New-Object System.Drawing.Size(160, 20)
    $comboAC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($option in $setting.PossibleValues) {
      $index = $comboAC.Items.Add($option.FriendlyName)
      if ($option.Index -eq $setting.CurrentACValue) {
        $comboAC.SelectedIndex = $index
      }
    }
    $editForm.Controls.Add($comboAC)

    $comboDC = New-Object System.Windows.Forms.ComboBox
    $comboDC.Location = New-Object System.Drawing.Point(160, 200)
    $comboDC.Size = New-Object System.Drawing.Size(160, 20)
    $comboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($option in $setting.PossibleValues) {
      $index = $comboDC.Items.Add($option.FriendlyName)
      if ($option.Index -eq $setting.CurrentDCValue) {
        $comboDC.SelectedIndex = $index
      }
    }
    $editForm.Controls.Add($comboDC)
  }

  $labelOtherPlans = New-Object System.Windows.Forms.Label
  $labelOtherPlans.Text = 'Same setting values for other power plan/s on this computer'
  $labelOtherPlans.Location = New-Object System.Drawing.Point(10, 240)
  $labelOtherPlans.Size = New-Object System.Drawing.Size(400, 20)
  $labelOtherPlans.Font = New-Object System.Drawing.Font('segoe ui', 8, [System.Drawing.FontStyle]::Bold)
  $editForm.Controls.Add($labelOtherPlans)

  $otherPlansGrid = New-Object System.Windows.Forms.DataGridView
  $otherPlansGrid.Location = New-Object System.Drawing.Point(10, 260)
  $otherPlansGrid.Size = New-Object System.Drawing.Size(400, 80)
  $otherPlansGrid.RowHeadersVisible = $false
  $otherPlansGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(219, 219, 219)
  $otherPlansGrid.EnableHeadersVisualStyles = $false
  $otherPlansGrid.GridColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
  $otherPlansGrid.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
  $otherPlansGrid.AllowUserToResizeRows = $false
  $otherPlansGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(160, 197, 231) #rgb(160, 197, 231)
  $otherPlansGrid.DefaultCellStyle.SelectionForeColor = 'Black'
  $otherPlansGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $otherPlansGrid.ReadOnly = $true
  $otherPlansGrid.AllowUserToAddRows = $false
  $otherPlansGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

 

  $otherPlansGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('segoe ui', 8, [System.Drawing.FontStyle]::Bold)

  $dataList = New-Object System.Collections.ArrayList

  $powerPlansRaw = powercfg /l
  $powerPlans = $powerPlansRaw | Where-Object { $_ -match '(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})\s+\((.*?)\)' } | ForEach-Object {
    $guid = $matches[1]
    $name = $matches[2]
    [PSCustomObject]@{ GUID = $guid; Name = "$name" }
  }
  #get same setting from other installed power plans
  foreach ($plan in $powerPlans) {
    if ($plan.GUID -eq $comboBox.SelectedItem.GUID) { continue } 

    $result = powercfg.exe /qh $plan.GUID $setting.SubgroupGUID $setting.PowerSettingGUID
    $acValue = $null
    $dcValue = $null

    foreach ($line in $result) {
      if ($line -like '*Current AC Power Setting Index*') {
        $splitLine = $line -split ':'
        $hexValue = $splitLine[1].Trim()
        $acValue = [Convert]::ToUInt32($hexValue, 16)
      }
      elseif ($line -like '*Current DC Power Setting Index*') {
        $splitLine = $line -split ':'
        $hexValue = $splitLine[1].Trim()
        $dcValue = [Convert]::ToUInt32($hexValue, 16)
      }
    }

    $acDisplay = $acValue
    $dcDisplay = $dcValue
    if ($setting.IsNumericRange) {
      $acDisplay = "$acValue $($setting.Units)"
      $dcDisplay = "$dcValue $($setting.Units)"
    }
    elseif ($setting.PossibleValues.Count -gt 0) {
      $acFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $acValue } | Select-Object -First 1).FriendlyName
      $dcFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $dcValue } | Select-Object -First 1).FriendlyName
      $acDisplay = if ($acFriendly) { $acFriendly } else { $acValue }
      $dcDisplay = if ($dcFriendly) { $dcFriendly } else { $dcValue }
    }

    $dataList.Add([PSCustomObject]@{
        'Power Plan'            = $plan.Name
        'Value AC (Plugged In)' = $acDisplay
        'Value DC (On Battery)' = $dcDisplay
      }) | Out-Null
  }

  $otherPlansGrid.DataSource = $dataList

  #unselect first row and set column back colors
  $otherPlansGrid.Add_DataBindingComplete({
      param($sender, $e)
      if ($sender.Columns['Value AC (Plugged In)'] -and $sender.Columns['Value DC (On Battery)']) {
        $acStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
        $acStyle.BackColor = [System.Drawing.Color]::FromArgb(193, 233, 233) 
        $sender.Columns['Value AC (Plugged In)'].DefaultCellStyle = $acStyle

        $dcStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
        $dcStyle.BackColor = [System.Drawing.Color]::FromArgb(205, 255, 200) 
        $sender.Columns['Value DC (On Battery)'].DefaultCellStyle = $dcStyle
      }
      #clear default selection
      $sender.ClearSelection()
      $sender.CurrentCell = $null
    })

  $editForm.Controls.Add($otherPlansGrid)

  $saveButton = New-Object System.Windows.Forms.Button
  $saveButton.Text = ' Save'
  $saveButton.Location = New-Object System.Drawing.Point(320, 350)
  $saveButton.Size = New-Object System.Drawing.Size(75, 30)
  $base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAs/AAALPwFJxTL7AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAARxJREFUOI2l078uhEEUBfDffLZkOx6AjUrhCTSyJBsKCQolEQ2JN1HoFDqNf42QrEqvVmCjUy3NZgvFJp/CLMN+uyRON2fuuefMzB3+idDDXBlVsoMFQSWyj3KXOvbVNPs3qFuWOZQb6WPYEqyrOuttULcsOC5M9R25YKXbJETxmExjgHNvko6KmmYWc2z/Iu4IDvAa1+V4T7JILA4Us6ZqC3cJv5g2GE82GtjEy6d4zom6PcwkdRUoFTi+GXak5daQCVXn6vYEu0XxugmeEm5K24Wyh1/EjTTBBaaTzVltN649Y6nIOWriM35MXwPlPsU/8eMZa5pyG8j/IM4F692Rzj7peadyq2gNdE6m8OsIKdLPxGSsuu/3mf6NdwQ5Tr+kaxstAAAAAElFTkSuQmCC'
  $imageBytes = [Convert]::FromBase64String($base64Icon)
  $memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
  $saveButton.Image = [System.Drawing.Image]::FromStream($memoryStream)
  $memoryStream.Dispose()
  $saveButton.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $saveButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
  $saveButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
  $saveButton.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
  $saveButton.Cursor = 'Hand'
  
  $saveButton.Add_Click({
      Update-Log -msg 'Saving Plan Setting(s)...' -type output

      if ($setting.IsNumericRange) {
        $setting.CurrentACValue = [int]($numericAC.Value)
        $setting.CurrentDCValue = [int]($numericDC.Value)
      }
      else {
        $selectedAC = $comboAC.SelectedItem
        $selectedDC = $comboDC.SelectedItem
        $acOption = $setting.PossibleValues | Where-Object { $_.FriendlyName -eq $selectedAC } | Select-Object -First 1
        $dcOption = $setting.PossibleValues | Where-Object { $_.FriendlyName -eq $selectedDC } | Select-Object -First 1
        $setting.CurrentACValue = $acOption.Index
        $setting.CurrentDCValue = $dcOption.Index
      }

      $planGuid = $comboBox.SelectedItem.GUID
      $subGuid = $setting.SubgroupGUID
      $setGuid = $setting.PowerSettingGUID

      #convert values to hex for powercfg
      $acHex = '0x' + $setting.CurrentACValue.ToString('X8')
      $dcHex = '0x' + $setting.CurrentDCValue.ToString('X8')
    
      powercfg /setacvalueindex $planGuid $subGuid $setGuid $acHex
      powercfg /setdcvalueindex $planGuid $subGuid $setGuid $dcHex

      Update-Display
      $editForm.Close()
    })
  $editForm.Controls.Add($saveButton)
 
  $editForm.ShowDialog()

}

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(1100, 520)
$searchBox.Size = New-Object System.Drawing.Size(160, 20)
$searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$searchBox.Text = 'Search...'
$searchBox.ForeColor = 'Gray'
$searchBox.add_Enter({
    if ($searchBox.Text -eq 'Search...') {
      $searchBox.Text = ''
      $searchBox.ForeColor = 'Black'
    }
  })
$searchBox.add_Leave({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
      $searchBox.Text = 'Search...'
      $searchBox.ForeColor = 'Gray'
    }
  })

$searchBox.add_TextChanged({
    $searchText = $searchBox.Text.Trim().ToLower()
    if ($searchText -ne 'search...') {

      $dataGridView.Rows.Clear()

      if (-not $global:Settings -or -not $global:Settings.Subgroups) {
        return
      }

      $selectedNode = $treeView.SelectedNode
      $subgroupSettings = @()

      if ($selectedNode -and $selectedNode.Tag) {
        #filter on just selected subgroup
        $selectedSubgroup = $selectedNode.Tag
        $subgroupSettings = $selectedSubgroup.Settings | Where-Object {
            ($_.PowerSettingName.ToLower() -like "*$searchText*") -or
            ($_.Description -and $_.Description.ToLower() -like "*$searchText*")
        }
      }
      else {
        #filter on all settings subgroup isnt selected for some reason
        foreach ($subgroup in $global:Settings.Subgroups) {
          $matchingSettings = $subgroup.Settings | Where-Object {
                ($_.PowerSettingName.ToLower() -like "*$searchText*") -or
                ($_.Description -and $_.Description.ToLower() -like "*$searchText*")
          }
          $subgroupSettings += $matchingSettings
        }
      }

      foreach ($setting in $subgroupSettings) {
        $minValue = 'N/A'
        $maxValue = 'N/A'
        $acValue = $setting.CurrentACValue
        $dcValue = $setting.CurrentDCValue

        if (!($setting.PossibleValues) -and $setting.IsNumericRange -eq $false) {
          $setting.IsNumericRange = $true
          $setting.MinPossible = 0
          $setting.MaxPossible = 255
        }

        if ($setting.IsNumericRange) {
          $minValue = "$($setting.MinPossible) $($setting.Units)"
          $maxValue = "$($setting.MaxPossible) $($setting.Units)"
          $acValue = "$($setting.CurrentACValue) $($setting.Units)"
          $dcValue = "$($setting.CurrentDCValue) $($setting.Units)"
        }
        elseif ($setting.PossibleValues.Count -gt 0) {
          $indices = $setting.PossibleValues | ForEach-Object { [int]$_.Index }
          $minIndex = $indices | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
          $maxIndex = $indices | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

          $minValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $minIndex } | Select-Object -First 1).FriendlyName
          $maxValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $maxIndex } | Select-Object -First 1).FriendlyName

          $acFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentACValue } | Select-Object -First 1).FriendlyName
          $dcFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentDCValue } | Select-Object -First 1).FriendlyName
          $acValue = if ($acFriendly) { $acFriendly } else { $setting.CurrentACValue }
          $dcValue = if ($dcFriendly) { $dcFriendly } else { $setting.CurrentDCValue }
        }

        $rowIndex = $dataGridView.Rows.Add(
          $editColumn.Image,
          $setting.PowerSettingName,
          $setting.Description,
          $acValue,
          $dcValue,
          $minValue,
          $maxValue
        )

        $dataGridView.Rows[$rowIndex].Tag = $setting
      }

      $dataGridView.ClearSelection()
    }
    
  })
$form.Controls.Add($searchBox)

$tooltip7 = New-Object System.Windows.Forms.ToolTip
$tooltip7.SetToolTip($searchBox, 'Search for Power Plan Setting')

$editSetting.Add_Click({
    $row = $dataGridView.SelectedRows[0]
    Edit-Setting -row $row.Tag
    
  })

$lookupSetting.Add_Click({
    $row = $dataGridView.SelectedRows[0]
    $setting = $row.Tag.PowerSettingName
    Start-Process "https://www.google.com/search?q=$setting"
  })

$dataGridView.add_MouseDown({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
      $hitTestInfo = $dataGridView.HitTest($e.X, $e.Y)
      if ($hitTestInfo.RowIndex -ge 0) {
        $dataGridView.ClearSelection()
        $dataGridView.Rows[$hitTestInfo.RowIndex].Selected = $true
        $contextMenu.Show($dataGridView, $e.Location)
      }
    }
  })

$dataGridView.Add_CellClick({
    param($sender, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
      #show edit form
      Edit-Setting -row $dataGridView.Rows[$e.RowIndex].Tag
    }
  })

$comboBox.Add_SelectedIndexChanged({
    Update-Display
  })

$refreshButton.Add_Click({
    Update-Display
  })

$treeView.Add_AfterSelect({
    $dataGridView.Rows.Clear()
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode -and $selectedNode.Tag) {
      foreach ($setting in $selectedNode.Tag.Settings) {
        $minValue = 'N/A'
        $maxValue = 'N/A'
        $acValue = $setting.CurrentACValue
        $dcValue = $setting.CurrentDCValue

        if (!($setting.PossibleValues) -and $setting.IsNumericRange -eq $false) {
          $setting.IsNumericRange = $true
          $setting.MinPossible = 0
          $setting.MaxPossible = 255
        }

        if ($setting.IsNumericRange) {
          $minValue = "$($setting.MinPossible) $($setting.Units)"
          $maxValue = "$($setting.MaxPossible) $($setting.Units)"
          $acValue = "$($setting.CurrentACValue) $($setting.Units)"
          $dcValue = "$($setting.CurrentDCValue) $($setting.Units)"
        }
        elseif ($setting.PossibleValues.Count -gt 0) {
          $indices = $setting.PossibleValues | ForEach-Object { [int]$_.Index }
          $minIndex = $indices | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
          $maxIndex = $indices | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

          $minValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $minIndex } | Select-Object -First 1).FriendlyName
          $maxValue = ($setting.PossibleValues | Where-Object { $_.Index -eq $maxIndex } | Select-Object -First 1).FriendlyName

          $acFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentACValue } | Select-Object -First 1).FriendlyName
          $dcFriendly = ($setting.PossibleValues | Where-Object { $_.Index -eq $setting.CurrentDCValue } | Select-Object -First 1).FriendlyName
          $acValue = if ($acFriendly) { $acFriendly } else { $setting.CurrentACValue }
          $dcValue = if ($dcFriendly) { $dcFriendly } else { $setting.CurrentDCValue }
        }

        $rowIndex = $dataGridView.Rows.Add(
          $editColumn.Image,
          $setting.PowerSettingName,
          $setting.Description,
          $acValue,
          $dcValue,
          $minValue,
          $maxValue
        )

        $dataGridView.Rows[$rowIndex].Tag = $setting
      }
    }
    $dataGridView.ClearSelection()
  })

$logbox = New-Object System.Windows.Forms.TextBox
$logbox.Location = New-Object System.Drawing.Point(1045, 605)
$logbox.Size = New-Object System.Drawing.Size(230, 40)
$logbox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$logbox.BackColor = [System.Drawing.Color]::FromArgb(219, 219, 219)
$logbox.ReadOnly = $true
$logbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logbox.WordWrap = $true
$logbox.Multiline = $true
$form.Controls.Add($logbox)

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = 'Log:'
$logLabel.Location = New-Object System.Drawing.Point(1005, 608)
$logLabel.Size = New-Object System.Drawing.Size(40, 20)
$form.Controls.Add($logLabel)

$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(1070, 523) 
$pictureBox.Size = New-Object System.Drawing.Size(20, 18) 
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$base64Icon = 'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAQlAAAEJQGmsd4JAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAARZJREFUOI2d078rhXEUx/GXmxTDNVyKUpKU/0MyWRXFILHKYmK2mEwGGY0mg8nCoBgMBiQGv7pXMRrkGr7Hryf3eeTU06lzvp/39/z4Pvy0PqzhAi+oYRfjKCmw6RDVcYkd7H+L7aHSSDyGN1xjKJOrYD0gB2jJistR6gN6cipcDch8NjETibkcMbTiHufZxGYAOgoAsBFnuz4CJanHVzz+AXAX/vOyEqpoRvcfAL3hq9+Dk1HWQoG4HFWeZhNtuMETBhuIm3zNava3AyPSHGqYkFr6sH5sh/hKzoscjSrqeMaxtLJ6fNXwK40ApI0sSS/uFmfYwrA0g8OALOdB8qwdRwFZ/C+kU9rEG6YKf9FfrCa1dIKBd5MKRCD5YQKsAAAAAElFTkSuQmCC'
$imageBytes = [Convert]::FromBase64String($base64Icon)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$pictureBox.image = [System.Drawing.Image]::FromStream($memoryStream)
$memoryStream.Dispose()
$form.Controls.Add($pictureBox)

function Update-Log {
  param([string]$msg,
    [ValidateSet('output', 'error')]
    $type
  )
  if ($type -eq 'output') {
    $logbox.AppendText("[+] $msg`r`n")
  }
  else {
    $logbox.AppendText("[!] $msg`r`n")
  }
   
}


Update-Display
$form.ShowDialog()
