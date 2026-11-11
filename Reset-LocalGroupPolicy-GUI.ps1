#Requires -Version 7.0
#Requires -RunAsAdministrator
<#
Run: pwsh.exe -STA -File .\Reset-LocalGroupPolicy-GUI.ps1
Scope: local Registry.pol files only. This is NOT a full Windows/security reset.
Backups contain policy files, not a system image. No automatic restart.
#>
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'Windows is required.' }
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell 7.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Launch with: pwsh.exe -STA -File .\Reset-LocalGroupPolicy-GUI.ps1'
}
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
# Strict binary parser: UTF-16 delimiters, NUL-terminated strings, binary DWORDs.
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
public class LocalPolEntry {
 public string Key {get;set;} public string Value {get;set;}
 public uint Type {get;set;} public string Data {get;set;}
}
public static class LocalPolReader {
 static void Expect(BinaryReader r, char c) {
  if(r.ReadUInt16() != (ushort)c) throw new InvalidDataException("Invalid policy delimiter.");
 }
 static string Str(BinaryReader r) {
  var s = new StringBuilder(); ushort c;
  while((c=r.ReadUInt16())!=0) s.Append((char)c);
  return s.ToString();
 }
 public static LocalPolEntry[] Parse(byte[] bytes) {
  var entries=new List<LocalPolEntry>();
  using(var r=new BinaryReader(new MemoryStream(bytes))) {
   if(r.ReadUInt32()!=0x67655250 || r.ReadUInt32()!=1)
    throw new InvalidDataException("Unrecognized policy header/version.");
   while(r.BaseStream.Position<r.BaseStream.Length) {
    Expect(r,'['); string key=Str(r); Expect(r,';');
    string value=Str(r); Expect(r,';'); uint type=r.ReadUInt32();
    Expect(r,';'); uint size=r.ReadUInt32(); Expect(r,';');
    if(size>Int32.MaxValue || size>r.BaseStream.Length-r.BaseStream.Position)
     throw new InvalidDataException("Invalid policy data length.");
    byte[] b=r.ReadBytes((int)size); Expect(r,']'); string data;
    if(type==1 || type==2 || type==7)
     data=Encoding.Unicode.GetString(b).TrimEnd('\0').Replace("\0"," | ");
    else if(type==4 && b.Length==4) data=BitConverter.ToUInt32(b,0).ToString();
    else if(type==11 && b.Length==8) data=BitConverter.ToUInt64(b,0).ToString();
    else data=BitConverter.ToString(b);
    entries.Add(new LocalPolEntry {Key=key,Value=value,Type=type,Data=data});
   }
  }
  return entries.ToArray();
 }
}
'@
$script:files = @()
$script:rows = @()
$script:backup = $null
$script:proc = $null
$base = Join-Path $env:windir 'System32'
$form = [Windows.Forms.Form]::new()
$form.Text = 'Local Group Policy - Preview and Reset'
$form.Size = [Drawing.Size]::new(1150,780)
$form.MinimumSize = [Drawing.Size]::new(900,650)
$form.StartPosition = 'CenterScreen'
$layout = [Windows.Forms.TableLayoutPanel]::new()
$layout.Dock = 'Fill'; $layout.ColumnCount=1; $layout.RowCount=5
foreach ($height in @(90,45,0,160,30)) {
    $style=[Windows.Forms.RowStyle]::new()
    if ($height -eq 0) { $style.SizeType='Percent'; $style.Height=100 }
    else { $style.SizeType='Absolute'; $style.Height=$height }
    [void]$layout.RowStyles.Add($style)
}
$form.Controls.Add($layout)
$intro=[Windows.Forms.Label]::new()
$intro.Dock='Fill'; $intro.Padding=[Windows.Forms.Padding]::new(10)
$intro.Text="Preview only until you click Reset. Lists raw registry keys, values and stored data from LOCAL Registry.pol files (not friendly policy names). Reset removes ALL listed files, including any empty files. Rows are informational, not individual selections.`r`nDomain/Intune policies may reapply. Security policy, firewall rules and other policy components are outside this reset. Removing files may not undo every applied value."
$layout.Controls.Add($intro,0,0)
$bar=[Windows.Forms.FlowLayoutPanel]::new(); $bar.Dock='Fill'
function New-Button($text) {
    $b=[Windows.Forms.Button]::new(); $b.Text=$text; $b.AutoSize=$true; $b.Height=30
    $bar.Controls.Add($b); return $b
}
$scan=New-Button 'Scan / Refresh'
$export=New-Button 'Export Preview CSV'
$reset=New-Button 'Back Up and Reset All Listed Files'
$open=New-Button 'Open Backup Folder'
$reset.Enabled=$false; $export.Enabled=$false; $open.Enabled=$false
$layout.Controls.Add($bar,0,1)
$grid=[Windows.Forms.DataGridView]::new()
$grid.Dock='Fill'; $grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false
$grid.AllowUserToDeleteRows=$false; $grid.AutoSizeColumnsMode='Fill'
$grid.SelectionMode='FullRowSelect'; $grid.RowHeadersVisible=$false
foreach ($name in @('File','Scope','Registry key','Value / instruction','Type','Stored data')) {
    [void]$grid.Columns.Add($name,$name)
}
$layout.Controls.Add($grid,0,2)
$log=[Windows.Forms.TextBox]::new(); $log.Dock='Fill'; $log.Multiline=$true
$log.ReadOnly=$true; $log.ScrollBars='Both'; $log.WordWrap=$false
$layout.Controls.Add($log,0,3)
$status=[Windows.Forms.Label]::new(); $status.Dock='Fill'; $status.Text='Click Scan / Refresh to preview.'
$layout.Controls.Add($status,0,4)
function Write-Log($message) {
    $line="$(Get-Date -Format 'HH:mm:ss')  $message"
    $log.AppendText($line+"`r`n")
    if ($script:backup) { Add-Content -LiteralPath (Join-Path $script:backup 'Activity.log') -Value $line }
}
function Scan-Policies {
    $reset.Enabled=$false; $export.Enabled=$false; $grid.Rows.Clear()
    $script:files=@(); $script:rows=@()
    try {
        foreach ($root in @('GroupPolicy','GroupPolicyUsers')) {
            $path=Join-Path $base $root
            if (-not (Test-Path -LiteralPath $path)) { continue }
            foreach ($f in Get-ChildItem -LiteralPath $path -Filter Registry.pol -Recurse -Force -File) {
                $bytes=[IO.File]::ReadAllBytes($f.FullName)
                $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
                $entries=@([LocalPolReader]::Parse($bytes))
                $relative=[IO.Path]::GetRelativePath($base,$f.FullName)
                $scope=if ($relative -match '^GroupPolicyUsers\\([^\\]+)') { "User SID: $($Matches[1])" }
                    elseif ($relative -match '\\Machine\\') { 'Computer (HKLM)' } else { 'User (HKCU)' }
                $script:files+= [pscustomobject]@{Path=$f.FullName; Relative=$relative; Hash=$hash}
                foreach ($e in $entries) {
                    $row=[pscustomobject]@{File=$relative; Scope=$scope; Key=$e.Key; Value=$e.Value; Type=$e.Type; Data=$e.Data}
                    $script:rows+=$row
                    [void]$grid.Rows.Add([object[]]@($relative,$scope,$e.Key,$e.Value,[string]$e.Type,$e.Data))
                }
                if ($entries.Count -eq 0) { [void]$grid.Rows.Add([object[]]@($relative,$scope,'(empty file)','','','')) }
                Write-Log "Scanned $relative : $($entries.Count) entries"
            }
        }
        $status.Text="$($script:files.Count) files; $($script:rows.Count) registry entries. Reset acts on every listed file."
        $reset.Enabled=$script:files.Count -gt 0; $export.Enabled=$script:rows.Count -gt 0
    } catch {
        $status.Text='Scan failed. Reset blocked; resolve the error and scan again.'
        Write-Log "SCAN ERROR: $_"
    }
}
$scan.Add_Click({ Scan-Policies })
$export.Add_Click({
    $dlg=[Windows.Forms.SaveFileDialog]::new(); $dlg.Filter='CSV files (*.csv)|*.csv'; $dlg.FileName='LocalPolicy-Preview.csv'
    try { if ($dlg.ShowDialog() -eq 'OK') { $script:rows | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding utf8; Write-Log "Exported $($dlg.FileName)" } }
    catch { [void][Windows.Forms.MessageBox]::Show("Export failed: $_") }
    finally { $dlg.Dispose() }
})
$open.Add_Click({ if ($script:backup) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:backup) } })
$timer=[Windows.Forms.Timer]::new(); $timer.Interval=500
$timer.Add_Tick({
    if ($script:proc -and $script:proc.HasExited) {
        $timer.Stop()
        $code=$script:proc.ExitCode
        $script:proc.Dispose(); $script:proc=$null
        try {
            Write-Log "gpupdate finished; exit code $code."
            foreach ($name in @('gpupdate-output.txt','gpupdate-error.txt')) {
                $p=Join-Path $script:backup $name
                if (Test-Path -LiteralPath $p) { $log.AppendText(([IO.File]::ReadAllText($p))+"`r`n") }
            }
        } catch { $log.AppendText("Log error: $_`r`n") }
        $status.Text=if ($code -eq 0) { 'Reset and refresh finished. Restart Windows, then verify settings.' } else { 'Files removed, but policy refresh failed. Review the log; restart and verify.' }
        $scan.Enabled=$true
    }
})
$reset.Add_Click({
    $answer=[Windows.Forms.MessageBox]::Show("Back up and remove ALL $($script:files.Count) listed local policy files, then run gpupdate /force?`r`n`r`nDomain/Intune policies can reapply. This does not reset every Windows security setting.",'Confirm local policy reset','YesNo','Warning','Button2')
    if ($answer -ne 'Yes') { return }
    $scan.Enabled=$false; $reset.Enabled=$false
    try {
        # Compare current contents with the reviewed scan before writing anything.
        foreach ($f in $script:files) {
            if ((Get-FileHash -LiteralPath $f.Path -Algorithm SHA256).Hash -ne $f.Hash) { throw "Policy changed since preview: $($f.Path). Scan again." }
        }
        $script:backup=Join-Path $env:ProgramData ('LocalPolicyBackup\'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        [void](New-Item -ItemType Directory -Path $script:backup)
        $open.Enabled=$true
        $script:files | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $script:backup 'Manifest.json')
        if ($script:rows.Count) { $script:rows | Export-Csv -LiteralPath (Join-Path $script:backup 'Preview.csv') -NoTypeInformation }
        # Verify ALL backup copies before the first deletion.
        foreach ($f in $script:files) {
            $dest=Join-Path $script:backup $f.Relative
            [void](New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force)
            Copy-Item -LiteralPath $f.Path -Destination $dest
            if ((Get-FileHash -LiteralPath $dest).Hash -ne $f.Hash) { throw "Backup verification failed: $($f.Path)" }
        }
        Write-Log "Verified backups: $script:backup"
        foreach ($f in $script:files) {
            if ((Get-FileHash -LiteralPath $f.Path).Hash -ne $f.Hash) { throw "Policy changed during reset: $($f.Path). Stopped; earlier deletions may have completed." }
            Remove-Item -LiteralPath $f.Path -Force
            Write-Log "REMOVED: $($f.Path)"
        }
        $status.Text='Files removed. Waiting for policy refresh...'
        $script:proc=Start-Process -FilePath (Join-Path $base 'gpupdate.exe') -ArgumentList '/force','/wait:60' -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $script:backup 'gpupdate-output.txt') -RedirectStandardError (Join-Path $script:backup 'gpupdate-error.txt')
        $timer.Start()
    } catch {
        $message="$_"
        try { Write-Log "RESET ERROR: $message" } catch { $log.AppendText("$message`r`n") }
        $status.Text='Reset stopped. Review log for completed actions; scan again before retrying.'
        $scan.Enabled=$true
        [void][Windows.Forms.MessageBox]::Show($message,'Reset stopped','OK','Error')
    }
})
$form.Add_FormClosing({ param($sender,$eventArgs)
    if ($script:proc -and -not $script:proc.HasExited) {
        $eventArgs.Cancel=$true
        [void][Windows.Forms.MessageBox]::Show('Wait for the policy refresh to finish before closing.')
    }
})
try { [void]$form.ShowDialog() } finally { $timer.Dispose(); $form.Dispose() }
