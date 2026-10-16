#Requires -Version 5.0
# ============================================================
#  批量重命名工具  BatchRename  v2.0
#  单文件 · 免安装 · 绿色版 · 支持高分屏（文字清晰）
#  实时预览 · 冲突检测 · 一键撤销
#  运行环境: Windows 7/8/10/11 + Windows PowerShell 5.1（系统自带）
# ============================================================

$ErrorActionPreference = 'Stop'

# ---------- 1. 声明 DPI 感知：让系统别拉伸界面（不然高分屏上文字发虚） ----------
$dpiApi = '[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value); ' +
          '[DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value); ' +
          '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();'
Add-Type -Namespace Brt -Name Dpi -MemberDefinition $dpiApi
$dpiOk = $false
try { $dpiOk = [Brt.Dpi]::SetProcessDpiAwarenessContext([IntPtr](-4)) } catch { }
if (-not $dpiOk) { try { $dpiOk = ([Brt.Dpi]::SetProcessDpiAwareness(2) -eq 0) } catch { } }
if (-not $dpiOk) { try { $dpiOk = [Brt.Dpi]::SetProcessDPIAware() } catch { } }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ---------- 2. 按屏幕 DPI 算缩放系数（96 DPI = 100%） ----------
$dpi = 96.0
try { $gg = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero); $dpi = [double]$gg.DpiX; $gg.Dispose() } catch { }
if ($dpi -lt 72 -or $dpi -gt 480) { $dpi = 96.0 }
$script:S = $dpi / 96.0

function U { param([double]$v) return [int][Math]::Round($v * $script:S) }

# ---------- 3. 配色 ----------
$C_Text   = [System.Drawing.Color]::FromArgb(32, 41, 55)
$C_Muted  = [System.Drawing.Color]::FromArgb(140, 148, 160)
$C_Dis    = [System.Drawing.Color]::FromArgb(178, 185, 196)
$C_Line   = [System.Drawing.Color]::FromArgb(230, 233, 238)
$C_White  = [System.Drawing.Color]::White
$C_Alt    = [System.Drawing.Color]::FromArgb(248, 250, 253)
$C_Accent = [System.Drawing.Color]::FromArgb(15, 108, 189)
$C_AccHov = [System.Drawing.Color]::FromArgb(12, 88, 155)
$C_Btn    = [System.Drawing.Color]::FromArgb(245, 246, 248)
$C_BtnHov = [System.Drawing.Color]::FromArgb(233, 236, 241)
$C_BtnBd  = [System.Drawing.Color]::FromArgb(212, 217, 224)
$C_OK     = [System.Drawing.Color]::FromArgb(0, 118, 62)
$C_Err    = [System.Drawing.Color]::FromArgb(190, 26, 26)
$C_ErrBg  = [System.Drawing.Color]::FromArgb(255, 240, 240)
$C_Warn   = [System.Drawing.Color]::FromArgb(176, 106, 0)

# ---------- 4. 字体（单位是点：DPI 感知后系统会自动放大像素，不用再乘缩放） ----------
$UiFont    = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)
$UiFontB   = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$FontSmall = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

$script:Files      = New-Object System.Collections.ArrayList
$script:UndoStack  = New-Object System.Collections.Stack
$script:MoveDone   = $null
$script:ErrorCount = 0

# ---------- 5. 控件工厂（所有尺寸统一走 U() 缩放） ----------
function New-Btn {
  param([string]$Text, [double]$X, [double]$Y, [double]$W, [double]$H, [switch]$Primary)
  $b = New-Object System.Windows.Forms.Button
  $b.Text = $Text
  $b.Font = if ($Primary) { $UiFontB } else { $UiFont }
  $b.SetBounds((U $X), (U $Y), (U $W), (U $H))
  $b.FlatStyle = 'Flat'
  $b.UseVisualStyleBackColor = $false
  $b.FlatAppearance.BorderSize = 1
  $b.Cursor = [System.Windows.Forms.Cursors]::Hand
  if ($Primary) {
    $b.BackColor = $C_Accent
    $b.ForeColor = $C_White
    $b.FlatAppearance.BorderColor = $C_Accent
    $b.FlatAppearance.MouseOverBackColor = $C_AccHov
    $b.FlatAppearance.MouseDownBackColor = $C_AccHov
  } else {
    $b.BackColor = $C_Btn
    $b.ForeColor = $C_Text
    $b.FlatAppearance.BorderColor = $C_BtnBd
    $b.FlatAppearance.MouseOverBackColor = $C_BtnHov
    $b.FlatAppearance.MouseDownBackColor = $C_BtnHov
  }
  return $b
}

function New-Lbl {
  param([string]$Text, [double]$X, [double]$Y, [double]$W, [double]$H, [switch]$Muted)
  $l = New-Object System.Windows.Forms.Label
  $l.Text = $Text
  $l.SetBounds((U $X), (U $Y), (U $W), (U $H))
  $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
  $l.Font = if ($Muted) { $FontSmall } else { $UiFont }
  $l.ForeColor = if ($Muted) { $C_Muted } else { $C_Text }
  return $l
}

function New-Rad {
  param([string]$Text, [double]$X, [double]$Y)
  $r = New-Object System.Windows.Forms.RadioButton
  $r.Text = $Text
  $r.Font = $UiFont
  $r.ForeColor = $C_Text
  $r.SetBounds((U $X), (U $Y), (U 152), (U 26))
  return $r
}

function New-Chk {
  param([string]$Text, [double]$X, [double]$Y, [double]$W = 120)
  $c = New-Object System.Windows.Forms.CheckBox
  $c.Text = $Text
  $c.Font = $UiFont
  $c.ForeColor = $C_Text
  $c.SetBounds((U $X), (U $Y), (U $W), (U 26))
  return $c
}

function New-Fld {
  param([string]$Kind, [double]$X, [double]$W)
  $c = New-Object ("System.Windows.Forms." + $Kind)
  $c.Font = $UiFont
  $c.Left = (U $X)
  $c.Width = (U $W)
  return $c
}

function Center-Row {
  param([double]$RowY, [double]$RowH, $Items)
  $top = (U $RowY)
  $band = (U $RowH)
  foreach ($c in $Items) {
    $c.Top = $top + [int][Math]::Round(($band - $c.Height) / 2)
  }
}

function New-HLine {
  param([double]$X, [double]$Y, [double]$W)
  $p = New-Object System.Windows.Forms.Panel
  $p.SetBounds((U $X), (U $Y), (U $W), 1)
  $p.BackColor = $C_Line
  return $p
}
# ============================================================
#                        核 心 逻 辑
# ============================================================
function Get-Rule {
  if ($rbFind.Checked)   { return 'find' }
  if ($rbAffix.Checked)  { return 'affix' }
  if ($rbNumber.Checked) { return 'number' }
  if ($rbCase.Checked)   { return 'case' }
  if ($rbExt.Checked)    { return 'ext' }
  return 'none'
}

function Get-NaturalKey {
  param([string]$s)
  if ([string]::IsNullOrEmpty($s)) { return '' }
  $parts = [regex]::Matches($s.ToLowerInvariant(), '\d+|\D+')
  $out = New-Object System.Collections.ArrayList
  foreach ($m in $parts) {
    if ($m.Value -match '^\d+$') { [void]$out.Add($m.Value.PadLeft(20, '0')) }
    else { [void]$out.Add($m.Value) }
  }
  return ($out -join '')
}

function Test-BadName {
  param([string]$n)
  if ([string]::IsNullOrWhiteSpace($n)) { return $true }
  if ($n -match '[\\/:*?"<>|]') { return $true }
  if ($n.EndsWith('.') -or $n.EndsWith(' ')) { return $true }
  if ($n -eq '.' -or $n -eq '..') { return $true }
  return $false
}

function Get-NewName {
  param($Item, [int]$Index)

  $full = $Item.Name
  $name = [System.IO.Path]::GetFileNameWithoutExtension($full)
  $ext  = [System.IO.Path]::GetExtension($full)
  $rule = Get-Rule
  if ($rule -eq 'none') { return $full }

  if ($rule -eq 'ext') {
    if ($chkExtDrop.Checked) { return $name }
    $e = $txtExt.Text.Trim().TrimStart('.')
    if ($e -eq '') { return $full }
    return ($name + '.' + $e)
  }

  $useExt = ($rule -ne 'number') -and $chkExt.Checked
  $target = if ($useExt) { $full } else { $name }
  $out = $target

  switch ($rule) {
    'find' {
      $f = $txtFind.Text
      if ($f -ne '') {
        $r = $txtReplace.Text
        if ($chkRegex.Checked) {
          $opt = if ($chkCase.Checked) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
          try { $out = [regex]::Replace($target, $f, $r, $opt) } catch { $out = $target }
        } else {
          $opt = if ($chkCase.Checked) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
          $out = [regex]::Replace($target, [regex]::Escape($f), $r.Replace('$', '$$'), $opt)
        }
      }
    }
    'affix' {
      $out = ($txtPrefix.Text + $target + $txtSuffix.Text)
    }
    'number' {
      $base = $txtBase.Text
      $dg   = [int]$numDigits.Value
      $st   = [int]$numStart.Value
      $sp   = [int]$numStep.Value
      $v    = $st + $Index * $sp
      $nf   = if ($v -lt 0) { '-' + ([Math]::Abs($v)).ToString('D' + $dg) } else { $v.ToString('D' + $dg) }
      $out  = if ($chkNumFirst.Checked) { $nf + $base } else { $base + $nf }
    }
    'case' {
      switch ($cmbCase.SelectedIndex) {
        0 { $out = $target.ToLower() }
        1 { $out = $target.ToUpper() }
        2 { if ($target.Length -gt 0) { $out = $target.Substring(0, 1).ToUpper() + $target.Substring(1) } }
        3 { $out = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($target.ToLower()) }
      }
    }
  }

  if ($useExt) { return $out }
  return ($out + $ext)
}

function Set-Status {
  param([string]$Text, [string]$Kind = '')
  $lblStatus.Text = $Text
  switch ($Kind) {
    'ok'    { $lblStatus.ForeColor = $C_OK }
    'err'   { $lblStatus.ForeColor = $C_Err }
    'warn'  { $lblStatus.ForeColor = $C_Warn }
    default { $lblStatus.ForeColor = $C_Text }
  }
  $lblStatus.Refresh()
}

function Update-RuleEnabled {
  $rule = Get-Rule
  foreach ($c in @($txtFind, $txtReplace, $chkCase, $chkRegex)) { $c.Enabled = ($rule -eq 'find') }
  foreach ($c in @($txtPrefix, $txtSuffix))                     { $c.Enabled = ($rule -eq 'affix') }
  foreach ($c in @($txtBase, $numStart, $numDigits, $numStep, $chkNumFirst)) { $c.Enabled = ($rule -eq 'number') }
  $cmbCase.Enabled = ($rule -eq 'case')
  $txtExt.Enabled  = (($rule -eq 'ext') -and (-not $chkExtDrop.Checked))
  $chkExtDrop.Enabled = ($rule -eq 'ext')
  $chkExt.Enabled  = ($rule -ne 'ext' -and $rule -ne 'number')
  if ($script:RowLabels) {
    foreach ($k in $script:RowLabels.Keys) {
      $on = ($rule -eq $k)
      foreach ($l in $script:RowLabels[$k]) {
        if ($l) { $l.ForeColor = if ($on) { $C_Text } else { $C_Dis } }
      }
    }
  }
}

function Update-Preview {
  if ($null -eq $lv) { return }
  $lv.BeginUpdate()
  try {
    $lv.Items.Clear()
    $n = $script:Files.Count
    $rows = New-Object System.Collections.ArrayList
    $changedOrig = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    for ($i = 0; $i -lt $n; $i++) {
      $it  = $script:Files[$i]
      $new = Get-NewName -Item $it -Index $i
      $dir = [System.IO.Path]::GetDirectoryName($it.FullName)
      if ([string]::IsNullOrEmpty($dir)) { $dir = '.' }
      $np  = Join-Path $dir $new
      $ch  = ($np -ne $it.FullName)
      if ($ch) { [void]$changedOrig.Add($it.FullName) }
      [void]$rows.Add([pscustomobject]@{ Item = $it; New = $new; NewPath = $np; Dir = $dir; Changed = $ch })
    }

    $seen = @{}
    $errCount = 0
    $chgCount = 0
    $idx = 0
    foreach ($r in $rows) {
      $err = ''
      if ($r.Changed) {
        $chgCount++
        if (Test-BadName $r.New) {
          $err = '名称非法'
        } else {
          $key = $r.NewPath.ToLowerInvariant()
          if ($seen.ContainsKey($key)) { $err = '名称重复' } else { $seen[$key] = $true }
          if ($err -eq '') {
            if ((Test-Path -LiteralPath $r.NewPath) -and (-not $changedOrig.Contains($r.NewPath))) { $err = '已存在同名文件' }
          }
        }
      }

      $li = New-Object System.Windows.Forms.ListViewItem($r.Item.Name)
      [void]$li.SubItems.Add($r.New)
      [void]$li.SubItems.Add($r.Dir)
      [void]$li.SubItems.Add($err)
      $li.ToolTipText = if ($err -eq '') { $r.Item.FullName } else { $err }

      if ($err -ne '') {
        $errCount++
        $li.ForeColor = $C_Err
        $li.BackColor = $C_ErrBg
      } elseif ($r.Changed) {
        $li.ForeColor = $C_OK
      } else {
        $li.ForeColor = $C_Muted
      }
      if ($err -eq '' -and (-not $r.Changed) -and ($idx % 2 -eq 1)) { $li.BackColor = $C_Alt }
      if ($err -eq '' -and $r.Changed -and ($idx % 2 -eq 1)) { $li.BackColor = $C_Alt }

      [void]$lv.Items.Add($li)
      $idx++
    }

    $script:ErrorCount = $errCount
    if ($n -eq 0) { $lblCount.Text = '还没有添加文件' }
    else { $lblCount.Text = "共 $n 个文件 · 将重命名 $chgCount 个" }
    $btnRename.Enabled = ($errCount -eq 0 -and $chgCount -gt 0)
    $btnUndo.Enabled = ($script:UndoStack.Count -gt 0)
    $lblCount.ForeColor = if ($errCount -gt 0) { $C_Err } else { $C_Muted }

    if ($n -eq 0) {
      Set-Status '先添加文件：点「添加文件 / 添加文件夹」，或把文件直接拖进下面的列表。'
    } elseif ($errCount -gt 0) {
      Set-Status "有 $errCount 个文件有问题（已标红），改好规则再执行。" 'err'
    } elseif ($chgCount -eq 0) {
      Set-Status '当前规则不会改变任何文件名。'
    } else {
      Set-Status "预览正常，共 $chgCount 个文件待重命名。" 'ok'
    }
  } finally {
    $lv.EndUpdate()
  }
}
# ============================================================
#                        文 件 列 表
# ============================================================
function Add-Paths {
  param([string[]]$Paths, [bool]$Recurse = $false)
  if ($null -eq $Paths) { return }
  $existing = @{}
  foreach ($it in $script:Files) { $existing[$it.FullName.ToLowerInvariant()] = $true }
  $added = 0
  foreach ($p in $Paths) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }
    if (Test-Path -LiteralPath $p -PathType Container) {
      $opt = if ($Recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
      $fs = @()
      try { $fs = [System.IO.Directory]::EnumerateFiles($p, '*', $opt) } catch { }
      foreach ($f in $fs) {
        $k = $f.ToLowerInvariant()
        if (-not $existing.ContainsKey($k)) {
          $existing[$k] = $true
          [void]$script:Files.Add([pscustomobject]@{ FullName = $f; Name = [System.IO.Path]::GetFileName($f) })
          $added++
        }
      }
    } elseif (Test-Path -LiteralPath $p -PathType Leaf) {
      $f = (Resolve-Path -LiteralPath $p).Path
      $k = $f.ToLowerInvariant()
      if (-not $existing.ContainsKey($k)) {
        $existing[$k] = $true
        [void]$script:Files.Add([pscustomobject]@{ FullName = $f; Name = [System.IO.Path]::GetFileName($f) })
        $added++
      }
    }
  }
  if ($added -gt 0) { Sort-Files }
  Set-Status "已添加 $added 个文件。" 'ok'
  Update-Preview
}

function Sort-Files {
  if ($script:Files.Count -lt 2) { return }
  $idx = $cmbSort.SelectedIndex
  $list = @()
  foreach ($it in $script:Files) {
    $t = [datetime]::MinValue; $s = 0
    try {
      $fi = New-Object System.IO.FileInfo($it.FullName)
      $t = $fi.LastWriteTime; $s = $fi.Length
    } catch { }
    $list += [pscustomobject]@{ Item = $it; N = (Get-NaturalKey $it.Name); T = $t; S = $s; E = ([System.IO.Path]::GetExtension($it.Name)) }
  }
  switch ($idx) {
    1 { $sorted = $list | Sort-Object T, N }
    2 { $sorted = $list | Sort-Object @{Expression='S';Descending=$true}, N }
    3 { $sorted = $list | Sort-Object E, N }
    default { $sorted = $list | Sort-Object N }
  }
  $script:Files = New-Object System.Collections.ArrayList
  foreach ($x in $sorted) { [void]$script:Files.Add($x.Item) }
}

# ============================================================
#                        执 行 / 撤 销
# ============================================================
function Sync-Items {
  param($Done)
  if ($Done.Count -eq 0) { return }
  $map = @{}
  foreach ($h in $Done) {
    $from = $h.From
    $map[$from] = $h.To
    foreach ($k in @($map.Keys)) {
      if ($k -ne $from -and $map[$k] -eq $from) { $map[$k] = $h.To }
    }
  }
  foreach ($it in $script:Files) {
    if ($map.ContainsKey($it.FullName)) {
      $p = $map[$it.FullName]
      $it.FullName = $p
      $it.Name = [System.IO.Path]::GetFileName($p)
    }
  }
}

function Invoke-MovePlan {
  param($Plan)
  $queue = New-Object System.Collections.ArrayList
  foreach ($p in $Plan) { [void]$queue.Add([pscustomobject]@{ From = $p.From; To = $p.To }) }
  $done = New-Object System.Collections.ArrayList
  $script:MoveDone = $done
  $guard = 0
  while ($queue.Count -gt 0) {
    if (++$guard -gt 50000) { throw '执行中断：检测到异常的循环依赖。' }
    $moved = $false
    for ($i = 0; $i -lt $queue.Count; $i++) {
      $cur = $queue[$i]
      $blocked = $false
      for ($j = 0; $j -lt $queue.Count; $j++) {
        if ($j -ne $i -and $queue[$j].From -eq $cur.To) { $blocked = $true; break }
      }
      if ($blocked) { continue }
      Move-Item -LiteralPath $cur.From -Destination $cur.To -ErrorAction Stop
      [void]$done.Add([pscustomobject]@{ From = $cur.From; To = $cur.To })
      $queue.RemoveAt($i)
      $i--
      $moved = $true
    }
    if ((-not $moved) -and $queue.Count -gt 0) {
      # 环形重命名：先挪到临时名，再改到目标名
      $cur = $queue[0]
      $d = [System.IO.Path]::GetDirectoryName($cur.From)
      $tmp = Join-Path $d ('__brt_tmp_' + [guid]::NewGuid().ToString('N').Substring(0, 10))
      Move-Item -LiteralPath $cur.From -Destination $tmp -ErrorAction Stop
      [void]$done.Add([pscustomobject]@{ From = $cur.From; To = $tmp })
      $cur.From = $tmp
    }
  }
  return $done
}

function Build-Plan {
  $plan = New-Object System.Collections.ArrayList
  for ($i = 0; $i -lt $script:Files.Count; $i++) {
    $it  = $script:Files[$i]
    $dir = [System.IO.Path]::GetDirectoryName($it.FullName)
    if ([string]::IsNullOrEmpty($dir)) { $dir = '.' }
    $np = Join-Path $dir (Get-NewName -Item $it -Index $i)
    if ($np -ne $it.FullName) { [void]$plan.Add([pscustomobject]@{ From = $it.FullName; To = $np }) }
  }
  return $plan
}

function Invoke-Rename {
  $plan = Build-Plan
  if ($plan.Count -eq 0) { Set-Status '没有需要重命名的文件。' 'warn'; return }
  $msg = ''; $kind = 'ok'
  $btnRename.Enabled = $false
  $btnUndo.Enabled = $false
  [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
  try {
    $done = Invoke-MovePlan -Plan $plan
    Sync-Items -Done $done
    $script:UndoStack.Push($done)
    $msg = "已重命名 $($done.Count) 个文件，可以点「撤销上次」还原。"
  } catch {
    $partial = $script:MoveDone
    if ($partial -and $partial.Count -gt 0) {
      Sync-Items -Done $partial
      $script:UndoStack.Push($partial)
      $msg = "重命名时出错（已完成 $($partial.Count) 个）：" + $_.Exception.Message
      $kind = 'err'
      [System.Windows.Forms.MessageBox]::Show("重命名过程中出错：`r`n`r`n" + $_.Exception.Message + "`r`n`r`n已经改好的 $($partial.Count) 个文件，可以点「撤销上次」还原。", '批量重命名工具', 'OK', 'Warning') | Out-Null
    } else {
      $msg = "重命名失败：" + $_.Exception.Message
      $kind = 'err'
      [System.Windows.Forms.MessageBox]::Show("重命名失败：`r`n`r`n" + $_.Exception.Message, '批量重命名工具', 'OK', 'Warning') | Out-Null
    }
  } finally {
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
  }
  Update-Preview
  if ($msg -ne '') { Set-Status $msg $kind }
}

function Invoke-Undo {
  if ($script:UndoStack.Count -eq 0) { Set-Status '没有可撤销的操作。' 'warn'; return }
  $hist = $script:UndoStack.Pop()
  if ($null -eq $hist -or $hist.Count -eq 0) { Set-Status '上次操作没有可撤销的内容。' 'warn'; Update-Preview; return }

  $plan = New-Object System.Collections.ArrayList
  for ($i = $hist.Count - 1; $i -ge 0; $i--) {
    [void]$plan.Add([pscustomobject]@{ From = $hist[$i].To; To = $hist[$i].From })
  }

  $msg = ''; $kind = 'ok'
  $btnRename.Enabled = $false
  $btnUndo.Enabled = $false
  [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
  try {
    $done = Invoke-MovePlan -Plan $plan
    Sync-Items -Done $done
    $msg = "已撤销，还原 $($done.Count) 个文件。"
  } catch {
    $partial = $script:MoveDone
    if ($partial -and $partial.Count -gt 0) { Sync-Items -Done $partial }
    $msg = "撤销失败：" + $_.Exception.Message
    $kind = 'err'
    [System.Windows.Forms.MessageBox]::Show("撤销过程中出错：`r`n`r`n" + $_.Exception.Message, '批量重命名工具', 'OK', 'Warning') | Out-Null
  } finally {
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
  }
  Update-Preview
  if ($msg -ne '') { Set-Status $msg $kind }
}

function Invoke-Safe {
  param([scriptblock]$Action)
  try { & $Action } catch { Set-Status ("出错：" + $_.Exception.Message) 'err' }
}
# ============================================================
#                          主 界 面
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = '批量重命名工具'
$form.BackColor = $C_White
$form.Font = $UiFont
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.MinimumSize = New-Object System.Drawing.Size((U 860), (U 520))

# 初始尺寸按屏幕可用区域自适应，避免在小屏 / 高分屏上超出屏幕
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$vw = [int](($wa.Width  - (U 60)) / $script:S)
$vh = [int](($wa.Height - (U 60)) / $script:S)
$cw = [Math]::Min(1000, [Math]::Max(880, $vw))
$ch = [Math]::Min(720, [Math]::Max(560, $vh))
$form.ClientSize = New-Object System.Drawing.Size((U $cw), (U $ch))
$form.StartPosition = 'CenterScreen'
$form.AllowDrop = $true

$tlp = New-Object System.Windows.Forms.TableLayoutPanel
$tlp.Dock = 'Fill'
$tlp.ColumnCount = 1
$tlp.RowCount = 3
[void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', (U 58))))
[void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', (U 58))))

# ---------------- 顶部工具条 ----------------
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Fill'
$topPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$topPanel.BackColor = $C_White

$btnAddFiles  = New-Btn '添加文件'   16 12 112 34
$btnAddFolder = New-Btn '添加文件夹' 138 12 126 34
$chkRecurse   = New-Chk '包含子文件夹' 276 16 140
$btnClear     = New-Btn '清空列表'   428 12 106 34
$lblCount     = New-Lbl '还没有添加文件' 560 12 420 34
$lblCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCount.ForeColor = $C_Muted

$lineTop = New-Object System.Windows.Forms.Panel
$lineTop.Dock = 'Bottom'
$lineTop.Height = 1
$lineTop.BackColor = $C_Line

[void]$topPanel.Controls.AddRange(@($btnAddFiles, $btnAddFolder, $chkRecurse, $btnClear, $lblCount, $lineTop))
$topPanel.Add_Resize({
  $w = $topPanel.ClientSize.Width - (U 588)
  if ($w -lt (U 120)) { $w = (U 120) }
  $lblCount.SetBounds((U 560), (U 12), $w, (U 34))
})

# ---------------- 规则区 ----------------
$midPanel = New-Object System.Windows.Forms.TableLayoutPanel
$midPanel.Dock = 'Fill'
$midPanel.ColumnCount = 1
$midPanel.RowCount = 2
[void]$midPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$midPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', (U 258))))
[void]$midPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = '重命名规则'
$grp.Font = $UiFontB
$grp.ForeColor = $C_Text
$grp.BackColor = $C_White
$grp.Dock = 'Fill'
$grp.Margin = New-Object System.Windows.Forms.Padding((U 14), (U 8), (U 14), (U 6))

$rbFind   = New-Rad '查找替换'       16 24
$lblFind  = New-Lbl '查找'           170 24 40 26
$txtFind  = New-Fld 'TextBox'        214 150
$lblRepl  = New-Lbl '替换为'         376 24 58 26
$txtReplace = New-Fld 'TextBox'      438 150
$chkCase  = New-Chk '区分大小写'     606 24 110
$chkRegex = New-Chk '正则'           720 24 74

$rbAffix  = New-Rad '添加 / 删除前后缀' 16 58
$lblPre   = New-Lbl '前缀'           170 58 40 26
$txtPrefix = New-Fld 'TextBox'       214 150
$lblSuf   = New-Lbl '后缀'           376 58 40 26
$txtSuffix = New-Fld 'TextBox'       420 150
$lblHint2 = New-Lbl '留空表示不加，可以只填一个' 588 58 270 26 -Muted

$rbNumber = New-Rad '批量编号'       16 92
$lblBase  = New-Lbl '基础名称'       170 92 68 26
$txtBase  = New-Fld 'TextBox'        242 140
$lblStart = New-Lbl '起始'           396 92 40 26
$numStart = New-Fld 'NumericUpDown'  440 66
$lblDig   = New-Lbl '位数'           518 92 40 26
$numDigits = New-Fld 'NumericUpDown' 562 66
$lblStep  = New-Lbl '步长'           640 92 40 26
$numStep  = New-Fld 'NumericUpDown'  684 66
$chkNumFirst = New-Chk '序号放前面'  764 92 130

$rbCase   = New-Rad '大小写转换'     16 126
$cmbCase  = New-Fld 'ComboBox'       170 280
$cmbCase.DropDownStyle = 'DropDownList'
[void]$cmbCase.Items.AddRange(@('全部小写：abc.txt', '全部大写：ABC.TXT', '首字母大写：Abc.txt', '每个单词首字母大写：Abc Def.txt'))

$rbExt    = New-Rad '修改扩展名'     16 160
$lblNewExt = New-Lbl '新扩展名'      170 160 68 26
$txtExt   = New-Fld 'TextBox'        242 150
$chkExtDrop = New-Chk '删除扩展名'   406 160 130
$lblHint5 = New-Lbl '不用输入点号，例如 jpg' 548 160 280 26 -Muted

$lineRule = New-HLine 16 196 900

$lblSort  = New-Lbl '列表排序'       16 206 68 28
$cmbSort  = New-Fld 'ComboBox'       88 160
$cmbSort.DropDownStyle = 'DropDownList'
[void]$cmbSort.Items.AddRange(@('文件名（自然顺序）', '修改时间（旧 → 新）', '文件大小（大 → 小）', '扩展名'))
$chkExt   = New-Chk '包含扩展名'     262 206 130 28
$lblHintOpt = New-Lbl '文件或文件夹可以直接拖进下面的列表' 410 206 350 28 -Muted
$btnReset = New-Btn '重置规则'       830 206 110 28

[void]$grp.Controls.AddRange(@(
  $rbFind, $lblFind, $txtFind, $lblRepl, $txtReplace, $chkCase, $chkRegex,
  $rbAffix, $lblPre, $txtPrefix, $lblSuf, $txtSuffix, $lblHint2,
  $rbNumber, $lblBase, $txtBase, $lblStart, $numStart, $lblDig, $numDigits, $lblStep, $numStep, $chkNumFirst,
  $rbCase, $cmbCase,
  $rbExt, $lblNewExt, $txtExt, $chkExtDrop, $lblHint5,
  $lineRule, $lblSort, $cmbSort, $chkExt, $lblHintOpt, $btnReset
))
$grp.Add_Resize({
  $w = $grp.ClientSize.Width - (U 34)
  if ($w -lt (U 200)) { $w = (U 200) }
  $lineRule.Width = $w
  $btnReset.Left = $grp.ClientSize.Width - $btnReset.Width - (U 18)
})

# ---------------- 文件列表 ----------------
$lv = New-Object System.Windows.Forms.ListView
$lv.Dock = 'Fill'
$lv.Margin = New-Object System.Windows.Forms.Padding((U 14), 0, (U 14), (U 12))
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.MultiSelect = $true
$lv.HideSelection = $false
$lv.GridLines = $false
$lv.BorderStyle = 'None'
$lv.HeaderStyle = 'Nonclickable'
$lv.AllowDrop = $true
$lv.ShowItemToolTips = $true
$lv.Font = $UiFont
[void]$lv.Columns.Add('原文件名', (U 330))
[void]$lv.Columns.Add('新文件名', (U 330))
[void]$lv.Columns.Add('所在文件夹', (U 200))
[void]$lv.Columns.Add('状态', (U 130))

$ctx = New-Object System.Windows.Forms.ContextMenuStrip
$ctx.Font = $UiFont
$miOpen = $ctx.Items.Add('在资源管理器中打开')
$miDel  = $ctx.Items.Add('从列表中移除')
$lv.ContextMenuStrip = $ctx

[void]$midPanel.Controls.Add($grp, 0, 0)
[void]$midPanel.Controls.Add($lv, 0, 1)

# ---------------- 底部操作条 ----------------
$botPanel = New-Object System.Windows.Forms.Panel
$botPanel.Dock = 'Fill'
$botPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$botPanel.BackColor = $C_White

$lineBot = New-Object System.Windows.Forms.Panel
$lineBot.Dock = 'Top'
$lineBot.Height = 1
$lineBot.BackColor = $C_Line

$btnRename = New-Btn '开始重命名' 16 12 150 36 -Primary
$btnUndo   = New-Btn '撤销上次'   178 12 116 36
$btnUndo.Enabled = $false
$lblStatus = New-Lbl '' 310 12 650 36

[void]$botPanel.Controls.AddRange(@($lineBot, $btnRename, $btnUndo, $lblStatus))
$botPanel.Add_Resize({
  $w = $botPanel.ClientSize.Width - (U 326)
  if ($w -lt (U 120)) { $w = (U 120) }
  $lblStatus.SetBounds((U 310), (U 12), $w, (U 36))
})

[void]$tlp.Controls.Add($topPanel, 0, 0)
[void]$tlp.Controls.Add($midPanel, 0, 1)
[void]$tlp.Controls.Add($botPanel, 0, 2)
[void]$form.Controls.Add($tlp)

# ---------------- 输入类控件垂直居中（等句柄创建后再校准一次） ----------------
$script:FieldRows = @(
  @{ Y = 24.0;  H = 26.0; Items = @($txtFind, $txtReplace) },
  @{ Y = 58.0;  H = 26.0; Items = @($txtPrefix, $txtSuffix) },
  @{ Y = 92.0;  H = 26.0; Items = @($txtBase, $numStart, $numDigits, $numStep) },
  @{ Y = 126.0; H = 26.0; Items = @($cmbCase) },
  @{ Y = 160.0; H = 26.0; Items = @($txtExt) },
  @{ Y = 206.0; H = 28.0; Items = @($cmbSort) }
)
function Align-Fields {
  foreach ($r in $script:FieldRows) { Center-Row $r.Y $r.H $r.Items }
}
Align-Fields

# ---------------- 悬停提示（替代满屏的小字说明） ----------------
$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 20000
$tip.InitialDelay = 400
$tip.ReshowDelay = 100
$tip.SetToolTip($txtFind, '要查找的文字；勾选「正则」后可用正则表达式')
$tip.SetToolTip($txtReplace, '替换成什么，留空表示把查找到的内容删掉')
$tip.SetToolTip($chkCase, '区分字母大小写')
$tip.SetToolTip($chkRegex, '把「查找」当成正则表达式处理')
$tip.SetToolTip($txtPrefix, '加在文件名最前面的内容')
$tip.SetToolTip($txtSuffix, '加在文件名最后面的内容')
$tip.SetToolTip($txtBase, '编号前面的固定文字，例如：照片_')
$tip.SetToolTip($numStart, '第一个文件的编号')
$tip.SetToolTip($numDigits, '编号的位数，3 表示 001、002 …')
$tip.SetToolTip($numStep, '每个文件编号增加多少，填 1 就是 1、2、3 …')
$tip.SetToolTip($chkNumFirst, '把序号放在基础名称前面')
$tip.SetToolTip($cmbCase, '选择大小写转换方式')
$tip.SetToolTip($txtExt, '新的扩展名，不用输入点号')
$tip.SetToolTip($chkExtDrop, '直接去掉文件的扩展名')
$tip.SetToolTip($chkExt, '让扩展名也参与改名（对查找替换 / 前后缀 / 大小写生效）')
$tip.SetToolTip($cmbSort, '列表顺序决定编号顺序；「自然顺序」会把 2 排在 10 前面')
$tip.SetToolTip($btnReset, '把所有规则恢复为默认状态')
$tip.SetToolTip($btnRename, '按上面预览的结果执行重命名')
$tip.SetToolTip($btnUndo, '撤销上一次重命名，可连续撤销')
$tip.SetToolTip($btnAddFiles, '选择一个或多个文件加入列表')
$tip.SetToolTip($btnAddFolder, '把某个文件夹里的文件全部加入列表')
$tip.SetToolTip($chkRecurse, '添加文件夹时，连同子文件夹里的文件一起加入')
$tip.SetToolTip($btnClear, '清空列表（不会删除文件）')

# ---------------- 规则行 ↔ 标签，用于把未选中的行整行变灰 ----------------
$script:RowLabels = @{
  'find'   = @($lblFind, $lblRepl)
  'affix'  = @($lblPre, $lblSuf, $lblHint2)
  'number' = @($lblBase, $lblStart, $lblDig, $lblStep)
  'case'   = @()
  'ext'    = @($lblNewExt, $lblHint5)
}
# ============================================================
#                          默 认 值
# ============================================================
$rbFind.Checked = $true
$cmbCase.SelectedIndex = 0
$cmbSort.SelectedIndex = 0
$numStart.Minimum = -999999; $numStart.Maximum = 999999; $numStart.Value = 1
$numDigits.Minimum = 1; $numDigits.Maximum = 12; $numDigits.Value = 3
$numStep.Minimum = -1000; $numStep.Maximum = 1000; $numStep.Value = 1

# ============================================================
#                            事 件
# ============================================================
$previewHandler = { Invoke-Safe { Update-RuleEnabled; Update-Preview } }

$ruleCtrls = @($txtFind, $txtReplace, $chkCase, $chkRegex, $txtPrefix, $txtSuffix,
               $txtBase, $numStart, $numDigits, $numStep, $chkNumFirst, $cmbCase,
               $txtExt, $chkExtDrop, $chkExt, $cmbSort,
               $rbFind, $rbAffix, $rbNumber, $rbCase, $rbExt)

foreach ($c in $ruleCtrls) {
  $c.Add_TextChanged($previewHandler)
  if ($c -is [System.Windows.Forms.RadioButton] -or $c -is [System.Windows.Forms.CheckBox]) {
    $c.Add_CheckedChanged($previewHandler)
  }
  if ($c -is [System.Windows.Forms.NumericUpDown]) { $c.Add_ValueChanged($previewHandler) }
  if ($c -is [System.Windows.Forms.ComboBox]) { $c.Add_SelectedIndexChanged($previewHandler) }
}

$btnAddFiles.Add_Click({
  Invoke-Safe {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '选择要重命名的文件（可多选）'
    $dlg.Multiselect = $true
    $dlg.Filter = '所有文件 (*.*)|*.*'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-Paths $dlg.FileNames $false }
  }
})

$btnAddFolder.Add_Click({
  Invoke-Safe {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择文件夹（其中的文件会被加入列表）'
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-Paths @($dlg.SelectedPath) $chkRecurse.Checked }
  }
})

$btnClear.Add_Click({ Invoke-Safe { $script:Files.Clear(); Update-Preview; Set-Status '列表已清空（文件没有被删除）。' 'warn' } })

$btnReset.Add_Click({
  Invoke-Safe {
    $rbFind.Checked = $true
    $txtFind.Text = ''; $txtReplace.Text = ''
    $chkCase.Checked = $false; $chkRegex.Checked = $false; $chkExt.Checked = $false
    $txtPrefix.Text = ''; $txtSuffix.Text = ''
    $txtBase.Text = ''; $numStart.Value = 1; $numDigits.Value = 3; $numStep.Value = 1; $chkNumFirst.Checked = $false
    $cmbCase.SelectedIndex = 0
    $txtExt.Text = ''; $chkExtDrop.Checked = $false
    $cmbSort.SelectedIndex = 0
    Sort-Files
    Update-RuleEnabled
    Update-Preview
    Set-Status '规则已重置。' 'warn'
  }
})

$btnRename.Add_Click({ Invoke-Safe { Invoke-Rename } })
$btnUndo.Add_Click({ Invoke-Safe { Invoke-Undo } })

$miOpen.Add_Click({
  Invoke-Safe {
    if ($lv.SelectedItems.Count -gt 0) {
      $name = $lv.SelectedItems[0].Text
      $p = $null
      foreach ($it in $script:Files) { if ($it.Name -eq $name) { $p = $it.FullName; break } }
      if (-not $p) { $p = Join-Path $lv.SelectedItems[0].SubItems[2].Text $name }
      if (Test-Path -LiteralPath $p) { Start-Process explorer.exe -ArgumentList ('/select,"' + $p + '"') }
      else { Start-Process explorer.exe -ArgumentList ([System.IO.Path]::GetDirectoryName($p)) }
    }
  }
})

$miDel.Add_Click({
  Invoke-Safe {
    $names = @()
    foreach ($si in $lv.SelectedItems) { $names += $si.Text }
    if ($names.Count -gt 0) {
      foreach ($n in $names) {
        for ($i = $script:Files.Count - 1; $i -ge 0; $i--) {
          if ($script:Files[$i].Name -eq $n) { $script:Files.RemoveAt($i); break }
        }
      }
      Update-Preview
    }
  }
})

$addDrop = {
  param($s, $e)
  if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
  } else {
    $e.Effect = [System.Windows.Forms.DragDropEffects]::None
  }
}
$doDrop = {
  param($s, $e)
  Invoke-Safe {
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Add-Paths $paths $chkRecurse.Checked
  }
}
$lv.Add_DragEnter($addDrop)
$lv.Add_DragDrop($doDrop)
$form.Add_DragEnter($addDrop)
$form.Add_DragDrop($doDrop)

$lv.Add_DoubleClick({
  Invoke-Safe { if ($lv.SelectedItems.Count -gt 0) { $miOpen.PerformClick() } }
})

$lv.Add_KeyDown({
  param($s, $e)
  if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) { Invoke-Safe { $miDel.PerformClick() } }
})

$lv.Add_Resize({
  Invoke-Safe {
    if ($lv.Columns.Count -lt 4) { return }
    $base = @((U 330), (U 330), (U 200), (U 130))
    $sum = $base[0] + $base[1] + $base[2] + $base[3]
    $avail = $lv.ClientSize.Width - 4
    if ($avail -lt (U 300)) { return }
    $k = 1.0
    if ($avail -lt $sum) { $k = [double]$avail / [double]$sum }
    $w = @()
    for ($i = 0; $i -lt 4; $i++) { $w += [int][Math]::Floor($base[$i] * $k) }
    $w[3] = $w[3] + ($avail - ($w[0] + $w[1] + $w[2] + $w[3]))
    if ($w[3] -lt (U 80)) { $w[3] = (U 80) }
    for ($i = 0; $i -lt 4; $i++) { $lv.Columns[$i].Width = $w[$i] }
  }
})

$form.Add_Shown({
  Invoke-Safe {
    Align-Fields
    Update-RuleEnabled
    Update-Preview
    Set-Status '就绪：先添加文件，选好规则并确认预览，然后点「开始重命名」。'
  }
})

# ============================================================
[System.Windows.Forms.Application]::Run($form)