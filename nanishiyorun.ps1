<#
MIT License

Copyright (c) 2022 masmt418

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script_base_name = (Get-Item $PSCommandPath).Basename
$config_name = $script_base_name + ".config.json"

$MUTEX_NAME = "Global\" + $script_base_name
#[System.Windows.Forms.MessageBox]::Show("" , "Title")

function main() {
  # duplicatehandle
  $mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME)
  if(-Not ($mutex.WaitOne(0, $false))) {
    $mutex.ReleaseMutex()
    $mutex.Close()
    return
  }

  # hide taskbar
  $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
  $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
  $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

  $application_context = New-Object System.Windows.Forms.ApplicationContext
  $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
  $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)

  # task tray
  $notify_icon = New-Object System.Windows.Forms.NotifyIcon
  $notify_icon.Icon = $icon
  $notify_icon.Visible = $true

  # context memu
  $notify_icon.ContextMenu = New-Object System.Windows.Forms.ContextMenu

  # config:breaks
  $breaks = @()
  Get-Content (".\" + $config_name) -Encoding UTF8 -Raw | ConvertFrom-Json | ForEach-Object { $_.breaks } | ForEach-Object {
    $property = [ordered]@{ start = $_.start; end =  $_.end; }
    $breaks += New-Object PsObject -Property $property
  }

  # event handler
  $item_Click = {
    ($sender, $e) = $this, $_

    try {

      DisableMenuItem $notify_icon.contextMenu.MenuItems

      $select_item = $sender.Text
      if($sender.Text -match "^\* (?<project>.+)$") {
        $select_item = $Matches.project
      }

      # get started MenuItem
      $started_items = @()
      [String[]]$started_items += GetStartedMenuItem $notify_icon.contextMenu.MenuItems

      # stop Csv
      if($started_items.Count -gt 0) {
        try {
          StopCsv $breaks $started_items
        } catch [System.IO.IOException] {
          NotifyError "Warning" $_.Exception $notify_icon
          return
        } catch [ApplicationException] {
          # cancel
          return
        } catch {
          NotifyError "Error" $_.Exception $notify_icon
          return
        }
      }

      # stop MenuItem
      StopMenuItem $notify_icon.contextMenu.MenuItems

      # start Csv
      if($started_items.Contains($select_item)) {
         return
      }
      try {
        StartCsv $select_item
      } catch [System.IO.IOException] {
        NotifyError "Warning" $_.Exception $notify_icon
        return
      } catch {
        return
      }

      # start MenuItem
      StartMenuItem $select_item $notify_icon.contextMenu.MenuItems

    } finally {
      EnableMenuItem $notify_icon.contextMenu.MenuItems
    }
  }

  $exit_Click = {
    ($sender, $e) = $this, $_

    try {

      DisableMenuItem $notify_icon.contextMenu.MenuItems

      # get started MenuItem
      $started_items = @()
      [String[]]$started_items += GetStartedMenuItem $notify_icon.contextMenu.MenuItems

      # stop Csv
      if($started_items.Count -gt 0) {
        try {
          StopCsv $breaks $started_items
        } catch [System.IO.IOException] {
          NotifyError "Warning" $_.Exception $notify_icon
          return
        } catch [ApplicationException] {
          # cancel
          return
        } catch {
          NotifyError "Error" $_.Exception $notify_icon
          return
        }
      }

      # stop MenuItems
      StopMenuItem $notify_icon.contextMenu.MenuItems

      $application_context.ExitThread()

    } finally {
      EnableMenuItem $notify_icon.contextMenu.MenuItems
    }
  }

  $notify_icon.add_MouseMove({
    ($sender, $e) = $this, $_
    if($e.Button -eq [Windows.Forms.MouseButtons]::Left) {
      $datetime = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
      $sender.Text = $datetime
    }
  })

  $notify_icon.add_DoubleClick({
    ($sender, $e) = $this, $_
    if($e.Button -eq [Windows.Forms.MouseButtons]::Left) {
      try {
        Start-Process (".\" + (Get-Date).ToString("yyyyMM") + ".csv")
      } catch {
      }
    }
  })

  # menu item (regist event)
  #https://docs.microsoft.com/ja-jp/dotnet/api/system.windows.forms.menu.menuitemcollection
  Get-Content (".\" + $config_name) -Encoding UTF8 -Raw | ConvertFrom-Json | ForEach-Object { $_.projects } | ForEach-Object {
    $notify_icon.contextMenu.MenuItems.Add((New-Object System.Windows.Forms.MenuItem -ArgumentList $_.project, $item_Click))
  }
  $notify_icon.contextMenu.MenuItems.Add((New-Object System.Windows.Forms.MenuItem -ArgumentList "-"))
  $notify_icon.contextMenu.MenuItems.Add((New-Object System.Windows.Forms.MenuItem -ArgumentList "Exit", $exit_Click))

  $started_items = @()
  [String[]]$started_items = GetStartedItemFromCsv
  $project = $_.project
  foreach($started_item in $started_items) {
    StartMenuItem $started_item $notify_icon.contextMenu.MenuItems
  }

  [void][System.Windows.Forms.Application]::Run($application_context)

  $notify_icon.Visible = $false
  $mutex.ReleaseMutex()
  $mutex.Close()
}

function StartMenuItem() {
  param(
    [System.String]$select_item,
    [System.Windows.Forms.MenuItem[]]$menu_items
  )

  foreach($item in $menu_items) {
    if($item.Text -eq $select_item) {
      # start item : * <project>
      $item.Text = "* " + $select_item
    }
  }
}

function StopMenuItem() {
  param(
    [System.Windows.Forms.MenuItem[]]$menu_items
  )
  foreach($item in $menu_items) {
    # started item : * <project>
    if($item.Text -match "^\* (?<project>.+)$") {
      $item.Text = $Matches.project
    }
  }
}

function GetStartedMenuItem() {
  param(
    [System.Windows.Forms.MenuItem[]]$menu_items
  )
  $started_items = @()
  foreach($item in $menu_items) {
    # started item : * <project>
    if($item.Text -match "^\* (?<project>.+)$") {
      $started_item = $Matches.project
      $started_items += $started_item
    }
  }

  # return
  $started_items
}

function GetStartedItemFromCsv() {
  $csv_datas = @()
  try {
    $csv_datas += Import-Csv (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
  } catch {
  }

  $started_items = @()
  foreach($csv_data in $csv_datas) {
    $allReadyExist = $false
    if($csv_data.end -eq "") {
      foreach($started_item in $started_items) {
        if($started_item -eq $csv_data.project) {
          $allReadyExist = $true
          continue
        }
      }
      if(-Not($allReadyExist)) {
        $started_items += $csv_data.project
      }
    }
  }

  # return
  $started_items 
}

function EnableMenuItem() {
  param(
    [System.Windows.Forms.MenuItem[]]$menu_items
  )
  foreach($item in $menu_items) {
    $item.Enabled = $true
  }
}

function DisableMenuItem() {
  param(
    [Parameter(Position=0)]
    [System.Windows.Forms.MenuItem[]]$menu_items
  )
  foreach($item in $menu_items) {
    $item.Enabled = $false
  }
}

function StartCsv() {
  param(
    [System.String]$select_item
  )

  $csv_datas = @()
  try {
    $csv_datas += Import-Csv (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
  } catch {
  }
  $property = [ordered]@{ project = "$select_item"; start = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss"); end = ""; description = "" }
  $new_row = New-Object PsObject -Property $property
  $csv_datas += $new_row
  $csv_datas | Export-Csv -NoTypeInformation (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
}

function StopCsv() {
  param(
    [PsCustomObject[]]$breaks,
    [String[]] $started_items
  )

  foreach($started_item in $started_items) {
    $csv_datas = @()
    $csv_datas_output = @()
    try {
      $csv_datas += Import-Csv (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
    } catch {
      $property = [ordered]@{ project = "$started_item"; start = ((Get-Date).ToString("yyyy/MM/dd") + " 00:00:00"); end = ""; description = "" }
      $csv_datas += New-Object PsObject -Property $property
    }
    foreach($csv_data in $csv_datas) {
      if(($csv_data.project -eq $started_item) -And ($csv_data.start -ne "") -And ($csv_data.end -eq "")) {
        # input description
        $description = ShowDescriptionDialog
        $end = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
        [PsCustomObject[]]$csv_datas_output += MakeCsvLine $csv_data.project $csv_data.start $end $breaks $description
      } else {
        $csv_datas_output += $csv_data
      }
    }
    $csv_datas_output | Export-Csv -NoTypeInformation (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
  }
}

function GetDescriptionInputHistoryFromCsv() {
  $csv_datas = @()
  try {
    $csv_datas += Import-Csv (".\" + (Get-Date).AddMonths(-1).ToString("yyyyMM") + ".csv") -Encoding Default
  } catch {
  }
  try {
    $csv_datas += Import-Csv (".\" + (Get-Date).ToString("yyyyMM") + ".csv") -Encoding Default
  } catch {
  }

  $descriptions = @()
  foreach($csv_data in $csv_datas) {
    $allReadyExist = $false
    if($csv_data.description -ne "") {
      foreach($description in $descriptions) {
        if($description -eq $csv_data.description) {
          $allReadyExist = $true
          continue
        }
      }
      if(-Not($allReadyExist)) {
        $descriptions += $csv_data.description
      }
    }
  }
  [array]::Reverse($descriptions)

  # return
  $descriptions
}

function ShowDescriptionDialog() {
  # description
  $form = New-Object System.Windows.Forms.Form
  #$form.Text = "description"
  $form.Size = New-Object System.Drawing.Size(310,35)
  $form.FormBorderStyle = "None"
  $form.Top = [System.Windows.Forms.Screen]::AllScreens.WorkingArea.Height - 35 - 100
  $form.Left = [System.Windows.Forms.Screen]::AllScreens.WorkingArea.Width - 310
  $form.StartPosition = "Manual"
  $form.Font = New-Object System.Drawing.Font("Meiryo UI",9)
  $comboBox = New-Object System.Windows.Forms.ComboBox
  $comboBox.Size = New-Object System.Drawing.Size(300,30)
  $comboBox.Location = New-Object System.Drawing.Point(5,10)
  $comboBox.DropDownStyle = "DropDown"
  $comboBox.FlatStyle = "standard"
  $comboBox.font = New-Object System.Drawing.Font("Meiryo UI",9)
  $comboBox.AutoCompleteMode = "SuggestAppend"
  [String[]]$descriptions = GetDescriptionInputHistoryFromCsv
  foreach($description in $descriptions) {
    [void]$comboBox.Items.Add($description)
  }
  $form.Controls.Add($comboBox)
  $okButton = New-Object System.Windows.Forms.Button
  $okButton.Size = New-Object System.Drawing.Size(20,20)
  $okButton.Location = New-Object System.Drawing.Point(310,10)
  $okButton.Text = "OK"
  $okButton.DialogResult = "OK"
  $form.Controls.Add($okButton)
  $cancelButton = New-Object System.Windows.Forms.Button
  $cancelButton.Size = New-Object System.Drawing.Size(20,20)
  $cancelButton.Location = New-Object System.Drawing.Point(310,10)
  $cancelButton.Text = "Cancel"
  $cancelButton.DialogResult = "Cancel"
  $form.Controls.Add($cancelButton)
  $form.AcceptButton = $okButton
  $form.CancelButton = $cancelButton
  $form.Add_Shown({$comboBox.Select()})
  $result = $form.ShowDialog()

  # return
  if ($result -eq "OK")
  {
    $comboBox.Text
  } elseif ($result -eq "Cancel") {
    throw [System.ApplicationException]::new()
  }
}

function MakeCsvLine
{
  param(
    [String]$project,
    [String]$start,
    [String]$end,
    [PsCustomObject[]]$breaks,
    [String]$description
  )

  # convert DateTime
  # $startDt
  $startDt = [DateTime]::ParseExact($start, "yyyy/MM/dd HH:mm:ss", $null)
  # $endDt
  #set-variable -name endDt -value [DateTime]::ParseExact($end, "yyyy/MM/dd HH:mm:ss", $null) -option constant
  $endDt = [DateTime]::ParseExact($end, "yyyy/MM/dd HH:mm:ss", $null)

  # $breaks
  $breakDts = @()
  [PsCustomObject[]]$breakDts = ConvertBreakDateTime $breaks $startDt $endDt

  # summarize by overlapping $breakDts --> $breakSumDts
  $breakSumDts = @()
  [PsCustomObject[]]$breakSumDts = SummarizeBreakDateTime $breakDts

  # round up/down $startDt $endDt
  foreach($breakDt in $breakSumDts) {
    # breakDt.startDt <= startDt <= breakDt.endDt --> breakDt.endDt
    if(($breakDt.startDt -le $startDt) -And ($startDt -le $breakDt.endDt)) {
      $startDt = $breakDt.endDt
    }
    # breakDt.startDt <= endDt <= breakDt.endDt --> breakDt.startDt
    if(($breakDt.startDt -le $endDt) -And ($endDt -le $breakDt.endDt)) {
      $endDt = $breakDt.startDt
    }
  }
  if($endDt -le $startDt) {
    return
  }

  $csv_datas = @()
  $baseDt = $startDt
  foreach($breakDt in $breakSumDts) {
    #_____|_______________|_______________|_____
    #   baseDt     breakDt.startDt      endDt --> add csv_data
    if(($baseDt -lt $breakDt.startDt) -And ($breakDt.startDt -lt $endDt)) {
      $csv_datas += SplitCsvLineByDate $project $baseDt $breakDt.startDt $description
      $baseDt = $breakDt.endDt
    }
  }
  # last line
  if($baseDt -lt $endDt) {
    $csv_datas += SplitCsvLineByDate $project $baseDt $endDt $description
  }

  # return
  $csv_datas
}

function ConvertBreakDateTime
{
  param(
    [PsCustomObject[]]$breaks,
    [DateTime]$startDt,
    [DateTime]$endDt
  )
  
  $breakDts = @()

  # $break convert DateTime --> $breakDts
  for($i = 0; $i -le $($endDt - $startDt).TotalDays; $i++) {
    foreach($break in $breaks) {
      $breakStartDt = [DateTime]::ParseExact(($startDt.AddDays($i).ToString("yyyy/MM/dd") + " " + $break.start.PadLeft(5, "0") + ":00"), "yyyy/MM/dd HH:mm:ss", $null)
      if($break.end.PadLeft(5, "0") -eq "00:00") {
        $breakEndDt = [DateTime]::ParseExact(($startDt.AddDays($i + 1).ToString("yyyy/MM/dd") + " " + $break.end.PadLeft(5, "0") + ":00"), "yyyy/MM/dd HH:mm:ss", $null)
      } else {
        $breakEndDt = [DateTime]::ParseExact(($startDt.AddDays($i).ToString("yyyy/MM/dd") + " " + $break.end.PadLeft(5, "0") + ":00"), "yyyy/MM/dd HH:mm:ss", $null)
      }
      $property = [ordered]@{ startDt = $breakStartDt; endDt = $breakEndDt; }
      $breakDts += New-Object PsObject -Property $property
    }
  }
  $breakDts = $breakDts | Sort-Object "startDt"
  
  # return
  $breakDts
}

function SummarizeBreakDateTime
{
  param(
    [PsCustomObject[]]$breakDts
  )

  $breakSumDts = @()

  # summarize by overlapping $breakDts --> $breakSumDts
  foreach($breakDt in $breakDts) {
    if($breakSumDts.Count -eq 0) {
      $breakSumDts += $breakDt
      continue
    }
    # breakDt.start <= previous breakDt.end
    if($breakDt.startDt -le $breakSumDts[$breakSumDts.Count - 1].endDt) {
      if($breakDt.endDt -le  $breakSumDts[$breakSumDts.Count - 1].endDt) {
        continue
      }
      $breakSumDts[$breakSumDts.Count - 1].endDt = $breakDt.endDt
    } else {
      $breakSumDts += $breakDt
    }
  }

  # return
  $breakSumDts
}

function SplitCsvLineByDate
{
  param(
    [String]$project,
    [DateTime]$startDt,
    [DateTime]$endDt,
    [String]$description
  )

  $csv_datas = @()

  # $endDt is the next day --> split data
  if($startDt.ToString("yyyy/MM/dd") -ne $endDt.ToString("yyyy/MM/dd")) {
    $property = [ordered]@{ project = $project; start = $startDt.ToString("yyyy/MM/dd HH:mm:ss"); end = ($endDt.AddDays(1).ToString("yyyy/MM/dd") + " 00:00:00"); description = $description }
    $csv_datas += New-Object PsObject -Property $property
    if("00:00:00" -ne $endDt.ToString("HH:mm:ss")) {
      $property = [ordered]@{ project = $project; start = ($endDt.ToString("yyyy/MM/dd") + " 00:00:00"); end = $endDt.ToString("yyyy/MM/dd HH:mm:ss"); description = $description }
      $csv_datas += New-Object PsObject -Property $property
    }
  } else {
    $property = [ordered]@{ project = $project; start = $startDt.ToString("yyyy/MM/dd HH:mm:ss"); end = $endDt.ToString("yyyy/MM/dd HH:mm:ss"); description = $description }
    $csv_datas += New-Object PsObject -Property $property
  }

  # return
  $csv_datas
}

function NotifyError()
{
  param(
    [Parameter(Position=0)]
    [ValidateSet("Warning", "Error")]
    [String]$type,
    [Parameter(Position=1)]
    [System.Exception]$exception,
    [Parameter(Position=2)]
    [System.Windows.Forms.NotifyIcon]$notify_icon
  )
  $notify_icon.BalloonTipIcon = $type
  $notify_icon.BalloonTipText = $exception.Message + " (" + "0x" + $exception.HResult.ToString("X8") + ")"
  $notify_icon.BalloonTipTitle = $script_base_name
  # This parameter is deprecated as of Windows Vista. Notification display times are now based on system accessibility settings.
  $notify_icon.ShowBalloonTip(1000)
}

main
