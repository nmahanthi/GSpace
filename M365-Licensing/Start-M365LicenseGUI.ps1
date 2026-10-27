<#
.SYNOPSIS
    Desktop window for Get-M365LicenseSavings.ps1 (Windows only).
.DESCRIPTION
    Fill in the tenant, tick the currencies, click Run. The window shows live progress, the key numbers,
    and the generated files (double-click a file to open it). It only calls the read-only licensing script.
    Sign-in happens in the normal Microsoft sign-in window.
.NOTES
    Requires Windows PowerShell 5.1 or PowerShell 7 on Windows, and: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Start it by double-clicking Start-M365LicenseGUI.cmd, or: powershell -STA -File .\Start-M365LicenseGUI.ps1
#>
if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) { throw 'This window needs Windows. On other systems run Get-M365LicenseSavings.ps1 directly.' }
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

# Use the module bundled with the app when present (offline install); otherwise fall back to the user's installed copy
$bundled = Join-Path $PSScriptRoot 'Modules'
if (Test-Path $bundled) { $env:PSModulePath = "$bundled$([IO.Path]::PathSeparator)$env:PSModulePath" }

$scriptPath = Join-Path $PSScriptRoot 'Get-M365LicenseSavings.ps1'
if (-not (Test-Path $scriptPath)) { [Windows.Forms.MessageBox]::Show("Cannot find $scriptPath", 'Licence savings') | Out-Null; return }

function New-Label($text, $x, $y, $w = 110) { $l = [Windows.Forms.Label]::new(); $l.Text = $text; $l.Location = [Drawing.Point]::new($x, $y + 3); $l.Size = [Drawing.Size]::new($w, 20); $l }
function New-Box($x, $y, $w, $text = '') { $t = [Windows.Forms.TextBox]::new(); $t.Location = [Drawing.Point]::new($x, $y); $t.Size = [Drawing.Size]::new($w, 24); $t.Text = $text; $t }
function New-Button($text, $x, $y, $w = 90) { $b = [Windows.Forms.Button]::new(); $b.Text = $text; $b.Location = [Drawing.Point]::new($x, $y - 1); $b.Size = [Drawing.Size]::new($w, 26); $b }

$form = [Windows.Forms.Form]::new()
$form.Text = 'Microsoft 365 Licence Savings'; $form.Size = [Drawing.Size]::new(940, 800); $form.StartPosition = 'CenterScreen'
$form.Font = [Drawing.Font]::new('Segoe UI', 9.5); $form.MinimumSize = [Drawing.Size]::new(940, 800)

# ---- Inputs ----------------------------------------------------------------
$form.Controls.Add((New-Label 'Tenant ID' 20 20))
$txtTenant = New-Box 140 20 360; $form.Controls.Add($txtTenant)
$lblHint = New-Label '(GUID or contoso.onmicrosoft.com)' 510 20 300; $lblHint.ForeColor = [Drawing.Color]::Gray; $form.Controls.Add($lblHint)

$form.Controls.Add((New-Label 'Customer name' 20 54))
$txtName = New-Box 140 54 360; $form.Controls.Add($txtName)
$form.Controls.Add((New-Label 'Optional. Blank = name from the tenant.' 510 54 300))

$form.Controls.Add((New-Label 'Currency' 20 88))
$cmbCur = [Windows.Forms.ComboBox]::new(); $cmbCur.DropDownStyle = 'DropDownList'; $cmbCur.Location = [Drawing.Point]::new(140, 86); $cmbCur.Size = [Drawing.Size]::new(150, 24)
$curCodes = @(); try { $curCodes = @((Import-Csv (Join-Path $PSScriptRoot 'prices.csv') | Select-Object -First 1).PSObject.Properties.Name | Where-Object { $_ -like 'Price*' -and $_.Length -eq 8 } | ForEach-Object { $_.Substring(5) }) } catch { $curCodes = @('USD') }
[void]$cmbCur.Items.Add('Auto (from tenant)'); foreach ($c in $curCodes) { [void]$cmbCur.Items.Add($c) }; if ($curCodes.Count -gt 1) { [void]$cmbCur.Items.Add('All') }
$cmbCur.SelectedIndex = 0; $form.Controls.Add($cmbCur)

$form.Controls.Add((New-Label 'Inactive after (days)' 310 88 130))
$numDays = [Windows.Forms.NumericUpDown]::new(); $numDays.Minimum = 30; $numDays.Maximum = 365; $numDays.Value = 90; $numDays.Location = [Drawing.Point]::new(445, 86); $numDays.Size = [Drawing.Size]::new(70, 24)
$form.Controls.Add($numDays)

$chkSkip = [Windows.Forms.CheckBox]::new(); $chkSkip.Text = 'Skip internal partner profit report'; $chkSkip.Location = [Drawing.Point]::new(140, 116); $chkSkip.Size = [Drawing.Size]::new(300, 24)
$form.Controls.Add($chkSkip)
if (-not (Test-Path (Join-Path $PSScriptRoot 'partner-terms.csv'))) { $chkSkip.Visible = $false }   # customer edition: no partner view

$form.Controls.Add((New-Label 'Prices file' 20 148))
$txtPrices = New-Box 140 148 520 (Join-Path $PSScriptRoot 'prices.csv'); $form.Controls.Add($txtPrices)
$btnPrices = New-Button 'Browse' 670 148 80; $btnEditPrices = New-Button 'Open' 760 148 80
$form.Controls.AddRange(@($btnPrices, $btnEditPrices))

$form.Controls.Add((New-Label 'Contract prices' 20 182))
$txtContract = New-Box 140 182 520 (Join-Path $PSScriptRoot 'customer-prices.csv'); $form.Controls.Add($txtContract)
$btnContract = New-Button 'Browse' 670 182 80; $btnEditContract = New-Button 'Open' 760 182 80
$form.Controls.AddRange(@($btnContract, $btnEditContract))
$lblContractHint = New-Label 'Optional: your own prices per licence. Overrides the list prices.' 140 208 700; $lblContractHint.ForeColor = [Drawing.Color]::Gray; $form.Controls.Add($lblContractHint)
$form.Controls.Add((New-Label 'Output folder' 20 236))
$txtOut = New-Box 140 236 520 (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'M365 Licence Review'); $form.Controls.Add($txtOut)
$btnOut = New-Button 'Browse' 670 236 80; $btnOpenOut = New-Button 'Open' 760 236 80
$form.Controls.AddRange(@($btnOut, $btnOpenOut))

$btnRun = New-Button 'Run analysis' 140 276 160; $btnRun.Height = 34; $btnRun.BackColor = [Drawing.Color]::FromArgb(10, 102, 194); $btnRun.ForeColor = [Drawing.Color]::White; $btnRun.FlatStyle = 'Flat'
$btnStop = New-Button 'Stop' 310 276 80; $btnStop.Enabled = $false
$btnStop.Add_Click({ $script:cancelled = $true; if ($script:ps) { try { $script:ps.Stop() } catch { } } })
$lblStatus = New-Label 'Ready.' 400 284 480
$form.Controls.AddRange(@($btnRun, $btnStop, $lblStatus))

# ---- Result cards ----------------------------------------------------------
$cards = @{}
$defs = @(@('spend', 'Monthly spend', 20), @('confirmed', 'Confirmed waste / month', 235), @('review', 'To review / month', 450), @('annual', 'Annual saving (confirmed)', 665))
foreach ($d in $defs) {
    $p = [Windows.Forms.Panel]::new(); $p.Location = [Drawing.Point]::new($d[2], 329); $p.Size = [Drawing.Size]::new(205, 78); $p.BorderStyle = 'FixedSingle'; $p.BackColor = [Drawing.Color]::FromArgb(243, 246, 250)
    $cap = [Windows.Forms.Label]::new(); $cap.Text = $d[1]; $cap.Location = [Drawing.Point]::new(8, 6); $cap.Size = [Drawing.Size]::new(190, 20); $cap.ForeColor = [Drawing.Color]::Gray
    $val = [Windows.Forms.Label]::new(); $val.Text = '-'; $val.Font = [Drawing.Font]::new('Segoe UI', 15, [Drawing.FontStyle]::Bold); $val.Location = [Drawing.Point]::new(8, 30); $val.Size = [Drawing.Size]::new(190, 36)
    $p.Controls.AddRange(@($cap, $val)); $form.Controls.Add($p); $cards[$d[0]] = $val
}
$cards.confirmed.ForeColor = [Drawing.Color]::FromArgb(192, 57, 43); $cards.annual.ForeColor = [Drawing.Color]::FromArgb(30, 158, 90)
$lblNote = New-Label '' 20 414 770; $lblNote.ForeColor = [Drawing.Color]::Gray; $form.Controls.Add($lblNote)
$cmbView = [Windows.Forms.ComboBox]::new(); $cmbView.DropDownStyle = 'DropDownList'; $cmbView.Location = [Drawing.Point]::new(800, 414); $cmbView.Size = [Drawing.Size]::new(100, 24); $cmbView.Anchor = 'Top,Right'; $cmbView.Visible = $false
$form.Controls.Add($cmbView)

# ---- Result tabs: licences, categories, findings, checks, log, files --------
function New-Grid([string[]]$cols) {
    $g = [Windows.Forms.DataGridView]::new(); $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false; $g.SelectionMode = 'FullRowSelect'; $g.BackgroundColor = [Drawing.Color]::White; $g.AutoSizeColumnsMode = 'Fill'
    foreach ($c in $cols) { [void]$g.Columns.Add($c, $c) }
    $g
}
function New-Tab($title, $control) { $t = [Windows.Forms.TabPage]::new(); $t.Text = $title; $t.Controls.Add($control); $t }
$gridSku  = New-Grid @('Licence', 'Type', 'Purchased', 'Assigned', 'Unassigned', 'Price / user / mo', 'Price source', 'Monthly spend', 'Unassigned cost / mo')
$gridProg = New-Grid @('Measure', 'Last run', 'Now', 'Change')
$gridCat  = New-Grid @('Category', 'Users', 'Saving / month')
$gridFind = New-Grid @('Category', 'User', 'Licences', 'Saving / month', 'Confidence', 'Recommended action')
$txtChecks = [Windows.Forms.TextBox]::new(); $txtChecks.Multiline = $true; $txtChecks.ReadOnly = $true; $txtChecks.ScrollBars = 'Vertical'; $txtChecks.Dock = 'Fill'; $txtChecks.BackColor = [Drawing.Color]::White
$txtLog = [Windows.Forms.TextBox]::new(); $txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.Dock = 'Fill'; $txtLog.Font = [Drawing.Font]::new('Consolas', 9)
$lstFiles = [Windows.Forms.ListBox]::new(); $lstFiles.Dock = 'Fill'
$tabs = [Windows.Forms.TabControl]::new(); $tabs.Location = [Drawing.Point]::new(20, 444); $tabs.Size = [Drawing.Size]::new(880, 290); $tabs.Anchor = 'Top,Left,Right,Bottom'
$tabLic = New-Tab 'Licences' $gridSku; $tabProg = New-Tab 'Progress' $gridProg; $tabCat = New-Tab 'By category' $gridCat; $tabFind = New-Tab 'Findings' $gridFind
$tabChk = New-Tab 'Checks' $txtChecks; $tabLog = New-Tab 'Log' $txtLog; $tabFiles = New-Tab 'Files (double-click to open)' $lstFiles
$tabs.TabPages.AddRange(@($tabLic, $tabProg, $tabCat, $tabFind, $tabChk, $tabLog, $tabFiles))
$form.Controls.Add($tabs)
$script:jsonByCur = @{}

function Add-LogLine([string]$msg) { $txtLog.AppendText($msg + [Environment]::NewLine); [Windows.Forms.Application]::DoEvents() }
function Set-Busy([bool]$busy) {
    foreach ($c in $form.Controls) { if ($c -is [Windows.Forms.TextBox] -and $c -ne $txtLog) { $c.Enabled = -not $busy } }
    $btnRun.Enabled = -not $busy; $form.UseWaitCursor = $busy
}
function Format-Amount([double]$v, [string]$cur) {
    $sym = switch ($cur) { 'INR' { [string][char]0x20B9 } 'USD' { '$' } 'GBP' { [string][char]0xA3 } 'EUR' { [string][char]0x20AC } default { "$cur " } }
    $culture = if ($cur -eq 'INR') { 'en-IN' } else { 'en-US' }
    $sym + $v.ToString('N0', [Globalization.CultureInfo]::GetCultureInfo($culture))
}

# ---- Buttons ---------------------------------------------------------------
$pickFile = { $d = [Windows.Forms.OpenFileDialog]::new(); $d.Filter = 'CSV files (*.csv)|*.csv'; if ($d.ShowDialog() -eq 'OK') { $txtPrices.Text = $d.FileName } }
$btnPrices.Add_Click($pickFile)
$btnContract.Add_Click({ $d = [Windows.Forms.OpenFileDialog]::new(); $d.Filter = 'CSV files (*.csv)|*.csv'; if ($d.ShowDialog() -eq 'OK') { $txtContract.Text = $d.FileName } })
$btnEditContract.Add_Click({ if (Test-Path $txtContract.Text) { Invoke-Item $txtContract.Text } else { [Windows.Forms.MessageBox]::Show('Contract prices file not found.', 'Licence savings') | Out-Null } })
$btnEditPrices.Add_Click({ if (Test-Path $txtPrices.Text) { Invoke-Item $txtPrices.Text } else { [Windows.Forms.MessageBox]::Show('Prices file not found.', 'Licence savings') | Out-Null } })
$btnOut.Add_Click({ $d = [Windows.Forms.FolderBrowserDialog]::new(); $d.SelectedPath = $txtOut.Text; if ($d.ShowDialog() -eq 'OK') { $txtOut.Text = $d.SelectedPath } })
$btnOpenOut.Add_Click({ if (Test-Path $txtOut.Text) { Invoke-Item $txtOut.Text } else { [Windows.Forms.MessageBox]::Show('Output folder does not exist yet. Run an analysis first.', 'Licence savings') | Out-Null } })
function Show-Results([string]$cur) {
    $jf = $script:jsonByCur[$cur]; if (-not $jf) { return }
    $r = Get-Content $jf -Raw | ConvertFrom-Json
    $cards.spend.Text = Format-Amount $r.MonthlySpend $cur
    $cards.confirmed.Text = Format-Amount $r.ConfirmedMonthlyWaste $cur
    $cards.review.Text = Format-Amount $r.ReviewMonthly $cur
    $cards.annual.Text = Format-Amount $r.ConservativeAnnualSaving $cur
    $gridSku.Rows.Clear(); $gridCat.Rows.Clear(); $gridFind.Rows.Clear(); $gridProg.Rows.Clear()
    $skus = @($r.Skus); $contract = 0; $trial = 0; $unpriced = 0
    foreach ($k in ($skus | Sort-Object { [double]$_.MonthlySpend } -Descending)) {
        $isTrial = ($k.Type -eq 'Trial/Free'); $hasPrice = ($null -ne $k.Price)
        if ($isTrial) { $trial++ } elseif (-not $hasPrice) { $unpriced++ }
        if ($k.PriceSource -eq 'Contract') { $contract++ }
        $price = if ($hasPrice) { Format-Amount ([double]$k.Price) $cur } else { 'n/a' }
        $src = switch ($k.PriceSource) { 'Contract' { 'Customer contract' } 'List' { 'List price' } default { 'No price' } }
        $idx = $gridSku.Rows.Add([object[]]@($k.Name, $(if ($isTrial) { 'Free / trial' } else { 'Paid' }), $k.Purchased, $k.Assigned, $k.Unassigned, $price, $src,
            (Format-Amount ([double]$k.MonthlySpend) $cur), (Format-Amount ([double]$k.UnassignedWaste) $cur)))
        if ($isTrial) { $gridSku.Rows[$idx].DefaultCellStyle.ForeColor = [Drawing.Color]::Gray }
        elseif (-not $hasPrice) { $gridSku.Rows[$idx].DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(255, 243, 205) }
        elseif ($k.PriceSource -eq 'Contract') { $gridSku.Rows[$idx].DefaultCellStyle.ForeColor = [Drawing.Color]::FromArgb(30, 158, 90) }
    }
    foreach ($c in @($r.Categories)) { [void]$gridCat.Rows.Add([object[]]@($c.Category, $c.Users, (Format-Amount ([double]$c.Monthly) $cur))) }
    $af = Join-Path (Split-Path $jf) ([IO.Path]::GetFileName($jf) -replace '-licensing-([A-Z]{3})\.json$', '-licensing-actions-$1.csv')
    $nFind = 0
    if (Test-Path $af) {
        foreach ($f in (Import-Csv $af | Select-Object -First 500)) {
            [void]$gridFind.Rows.Add([object[]]@($f.Category, "$($f.User) ($($f.UPN))", $f.Licences, (Format-Amount ([double]$f.MonthlySaving) $cur), $f.Confidence, $f.Action)); $nFind++
        }
    }
    $tabFind.Text = "Findings ($nFind)"
    # Progress since the previous run for this customer
    $h = $r.History; $progNote = ''
    if ($h) {
        $mm = @(@('Monthly licence spend', $h.SpendBefore, $h.SpendNow), @('Confirmed waste / month', $h.ConfirmedBefore, $h.ConfirmedNow),
                @('To review / month', $h.ReviewBefore, $h.ReviewNow), @('Unassigned seats cost / month', $h.UnassignedBefore, $h.UnassignedNow))
        foreach ($m in $mm) {
            $d = [double]$m[2] - [double]$m[1]; $sign = if ($d -gt 0) { '+' } else { '' }
            $i = $gridProg.Rows.Add([object[]]@($m[0], (Format-Amount ([double]$m[1]) $cur), (Format-Amount ([double]$m[2]) $cur), ($sign + (Format-Amount $d $cur))))
            if ([math]::Abs($d) -ge 0.5) { $gridProg.Rows[$i].Cells[3].Style.ForeColor = $(if ($d -lt 0) { [Drawing.Color]::FromArgb(30, 158, 90) } else { [Drawing.Color]::FromArgb(192, 57, 43) }) }
        }
        [void]$gridProg.Rows.Add([object[]]@('Seats removed since last run', '', '', "$($h.SeatsRemoved) (worth $(Format-Amount ([double]$h.SeatsRemovedMonthly) $cur)/month)"))
        [void]$gridProg.Rows.Add([object[]]@('Earlier findings no longer applying', '', '', "$($h.ResolvedCount) (worth $(Format-Amount ([double]$h.ResolvedMonthly) $cur)/month)"))
        [void]$gridProg.Rows.Add([object[]]@('New findings', '', '', "$($h.NewCount) (worth $(Format-Amount ([double]$h.NewMonthly) $cur)/month)"))
        if ($h.PriceChanged) { [void]$gridProg.Rows.Add([object[]]@('Note', '', '', 'Prices changed since the last run, so the spend change includes price changes.')) }
        $tabProg.Text = "Progress (vs $($h.PreviousRun))"
        if ($h.SeatsRemoved -gt 0) { $progNote = "Since $($h.PreviousRun): $($h.SeatsRemoved) seat(s) removed, about $(Format-Amount ([double]$h.SeatsRemovedMonthly) $cur)/month saved." }
    } else {
        [void]$gridProg.Rows.Add([object[]]@('First run for this customer', '', '', 'Baseline saved. The next run will show what changed.'))
        $tabProg.Text = 'Progress (baseline saved)'
    }
    $w = @($r.Warnings | Sort-Object { if ($_ -like 'FIX*') { 0 } else { 1 } })
    $txtChecks.Text = if ($w.Count) { ($w | ForEach-Object { '- ' + $_ }) -join ([Environment]::NewLine + [Environment]::NewLine) } else { 'No check warnings.' }
    $tabChk.Text = "Checks ($($w.Count))"
    $parts = @("Figures in $cur.", "Annual saving if every review item is accepted: $(Format-Amount $r.AnnualSaving $cur).")
    $parts += $(if ($contract) { "$contract licence price(s) from the customer's contract." } else { 'Using list prices (no contract prices applied).' })
    if ($trial) { $parts += "$trial free/trial product(s) not counted as spend." }
    if ($unpriced) { $parts += "$unpriced paid licence(s) have no price (yellow rows): fill them in under Contract prices." }
    if ($r.UsageNamesConcealed) { $parts += 'User names are hidden in usage reports, so the email-only check was skipped: see the first item under Checks for the fix.' }
    if ($progNote) { $parts += $progNote }
    $lblNote.Text = $parts -join ' '
    $lblNote.ForeColor = if ($unpriced -or $w.Count -or $r.UsageNamesConcealed) { [Drawing.Color]::FromArgb(176, 112, 0) } else { [Drawing.Color]::Gray }
}
$cmbView.Add_SelectedIndexChanged({ if ($cmbView.SelectedItem) { Show-Results ([string]$cmbView.SelectedItem) } })
$lstFiles.Add_DoubleClick({ if ($lstFiles.SelectedItem) { $f = Join-Path $txtOut.Text $lstFiles.SelectedItem; if (Test-Path $f) { Invoke-Item $f } } })

$btnRun.Add_Click({
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $ans = [Windows.Forms.MessageBox]::Show("The Microsoft Graph PowerShell module is needed (one-time, installs for your user only). Install it now?", 'Licence savings', 'YesNo', 'Question')
        if ($ans -ne 'Yes') { return }
        $lblStatus.Text = 'Installing Microsoft Graph module. This can take a few minutes.'; Set-Busy $true
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
        } catch { Set-Busy $false; [Windows.Forms.MessageBox]::Show("Install failed: $($_.Exception.Message)", 'Licence savings') | Out-Null; $lblStatus.Text = 'Module install failed.'; return }
        Set-Busy $false
    }
    $sel = [string]$cmbCur.SelectedItem
    $currencies = if ($sel -like 'Auto*') { @() } elseif ($sel -eq 'All') { $curCodes } else { @($sel) }
    if (-not (Test-Path $txtPrices.Text)) { [Windows.Forms.MessageBox]::Show('Prices file not found.', 'Licence savings') | Out-Null; return }

    $params = @{ InactiveDays = [int]$numDays.Value; PricesCsv = $txtPrices.Text; OutputDir = $txtOut.Text; Verbose = $true }
    if ($currencies.Count) { $params.Currencies = $currencies }
    if ($txtTenant.Text.Trim()) { $params.TenantId = $txtTenant.Text.Trim() }
    if ($txtName.Text.Trim()) { $params.CustomerName = $txtName.Text.Trim() }
    if ($txtContract.Text.Trim() -and (Test-Path $txtContract.Text)) { $params.CustomerPricesCsv = $txtContract.Text }
    if ($chkSkip.Checked) { $params.SkipPartnerReport = $true }

    $txtLog.Clear(); $lstFiles.Items.Clear(); foreach ($k in $cards.Keys) { $cards[$k].Text = '-' }; $lblNote.Text = ''
    $gridSku.Rows.Clear(); $gridCat.Rows.Clear(); $gridFind.Rows.Clear(); $gridProg.Rows.Clear(); $txtChecks.Clear(); $cmbView.Visible = $false; $cmbView.Items.Clear(); $script:jsonByCur = @{}
    $tabFind.Text = 'Findings'; $tabProg.Text = 'Progress'; $tabChk.Text = 'Checks'; $lblNote.ForeColor = [Drawing.Color]::Gray
    $start = Get-Date; $failed = $false
    Set-Busy $true; $lblStatus.Text = 'Running. A Microsoft sign-in window may open behind this one.'
    # Run in a background runspace so the window stays responsive (Graph calls can take minutes)
    $script:cancelled = $false
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({ param($sp, $pr) & $sp @pr }).AddArgument($scriptPath).AddArgument($params)
    $script:ps = $ps
    $btnStop.Enabled = $true
    $tick = Get-Date
    try {
        $async = $ps.BeginInvoke()
        while (-not $async.IsCompleted) {
            foreach ($m in $ps.Streams.Verbose.ReadAll())     { Add-LogLine ($m.Message) }
            foreach ($m in $ps.Streams.Warning.ReadAll())     { Add-LogLine ('WARNING: ' + $m.Message) }
            foreach ($m in $ps.Streams.Information.ReadAll()) { Add-LogLine ([string]$m.MessageData) }
            foreach ($m in $ps.Streams.Error.ReadAll())       { Add-LogLine ('ERROR: ' + $m.ToString()) }
            $lblStatus.Text = 'Running ({0:mm\:ss}). A Microsoft sign-in window may open behind this one.' -f ((Get-Date) - $start)
            [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 150
        }
        foreach ($m in $ps.Streams.Verbose.ReadAll())     { Add-LogLine ($m.Message) }
        foreach ($m in $ps.Streams.Warning.ReadAll())     { Add-LogLine ('WARNING: ' + $m.Message) }
        foreach ($m in $ps.Streams.Information.ReadAll()) { Add-LogLine ([string]$m.MessageData) }
        foreach ($m in $ps.Streams.Error.ReadAll())       { Add-LogLine ('ERROR: ' + $m.ToString()) }
        if ($script:cancelled) { $failed = $true; Add-LogLine 'Stopped by user.' }
        else { try { [void]$ps.EndInvoke($async) } catch { $failed = $true; Add-LogLine ('ERROR: ' + $_.Exception.Message) } }
    } catch { $failed = $true; Add-LogLine ('ERROR: ' + $_.Exception.Message) }
    finally { $btnStop.Enabled = $false; $ps.Dispose(); $script:ps = $null; try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }; Set-Busy $false }

    $out = $txtOut.Text
    $files = @(); if (Test-Path $out) { $files = @(Get-ChildItem $out -File | Where-Object { $_.LastWriteTime -ge $start.AddSeconds(-2) } | Sort-Object Name) }
    foreach ($f in $files) { $lstFiles.Items.Add($f.Name) | Out-Null }

    foreach ($f in ($files | Where-Object { $_.Name -match '-licensing-[A-Z]{3}\.json$' })) {
        $script:jsonByCur[[regex]::Match($f.Name, '-licensing-([A-Z]{3})\.json$').Groups[1].Value] = $f.FullName
    }
    $json = $script:jsonByCur.Count -gt 0
    if ($json) {
        foreach ($c in ($script:jsonByCur.Keys | Sort-Object)) { [void]$cmbView.Items.Add($c) }
        $cmbView.Visible = ($cmbView.Items.Count -gt 1)
        $cmbView.SelectedIndex = 0   # shows the first currency
        $tabs.SelectedTab = $tabLic
    } else { $tabs.SelectedTab = $tabLog }
    if ($failed -or -not $json) { $lblStatus.Text = 'Finished with errors. See the log.' } else { $lblStatus.Text = "Done. $($files.Count) file(s) written to $out" }
})

[void]$form.ShowDialog()
