Attribute VB_Name = "Audit_Security"
Sub SecurityPatchAndOperationsAudit()
    Dim wshShell As Object
    Dim execProcess As Object
    Dim psCommand As String
    Dim resultOutput As String
    Dim lines() As String
    Dim lineItem As Variant
    Dim rowNum As Long
    Dim currentSheet As Worksheet
    
    Set currentSheet = ActiveSheet
    currentSheet.Cells.Clear
    
    'Report
    With currentSheet
        .Cells(1, 1).Value = "ENTERPRISE CYBERSECURITY DAILY AUTOMATED COMPLIANCE REPORT"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Execution Timestamp: " & Now & " | Target Scope: Local Endpoint Baseline"
        .Cells(2, 1).Font.Italic = True
        
        .Cells(4, 1).Value = "Audit Objective / Component"
        .Cells(4, 2).Value = "Compliance Metric Status"
        .Cells(4, 3).Value = "Risk Severity Level"
        .Range("A4:C4").Font.Bold = True
        .Range("A4:C4").Interior.Color = RGB(220, 220, 220)
    End With

    psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """
    psCommand = psCommand & "$ErrorActionPreference = 'SilentlyContinue'; "
    
    'Windows Update Agent Patch Verification
    psCommand = psCommand & "Write-Output '--- SYSTEM PATCH COMPLIANCE MATRIX ---'; "
    psCommand = psCommand & "$updateSession = New-Object -ComObject 'Microsoft.Update.Session'; "
    psCommand = psCommand & "$updateSearcher = $updateSession.CreateUpdateSearcher(); "
    psCommand = psCommand & "$searchResult = $updateSearcher.Search('IsInstalled=0 and Type=''Software'''); "
    psCommand = psCommand & "if ($searchResult.Updates.Count -eq 0) { "
    psCommand = psCommand & "  Write-Output 'Operating System Patch Level: All Critical Updates Installed|COMPLIANT|LOW'; "
    psCommand = psCommand & "} else { "
    psCommand = psCommand & "  foreach ($update in $searchResult.Updates) { "
    psCommand = psCommand & "    $severity = 'MEDIUM'; "
    psCommand = psCommand & "    if ($update.MsrcSeverity -eq 'Critical' -or $update.Title -like '*Security*') { $severity = 'CRITICAL' }; "
    psCommand = psCommand & "    Write-Output ('Missing Update: ' + $update.Title.Replace('|','-') + '|NON-COMPLIANT|' + $severity); "
    psCommand = psCommand & "  } "
    psCommand = psCommand & "}; "
    
    'Local Administrators Account Review
    psCommand = psCommand & "Write-Output '--- IDENTITY GOVERNANCE & ACCESS CONTROLS ---'; "
    psCommand = psCommand & "$localAdmins = Net LocalGroup Administrators | Where-Object { $_ -and $_ -notlike '*The command*' -and $_ -notlike '*Members*' -and $_ -notlike '---*' }; "
    psCommand = psCommand & "foreach ($admin in $localAdmins) { "
    psCommand = psCommand & "  if ($admin.Trim() -and $admin.Trim() -ne 'Administrator') { "
    psCommand = psCommand & "    Write-Output ('Privileged Account Detected: ' + $admin.Trim() + '|REVIEW REQUIRED|MEDIUM'); "
    psCommand = psCommand & "  } "
    psCommand = psCommand & "}; "
    
    'Antivirus and Firewall Endpoint Baseline Verification
    psCommand = psCommand & "Write-Output '--- HOST PROTECTION ENDPOINT BASELINES ---'; "
    psCommand = psCommand & "$avStatus = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName 'AntivirusProduct'; "
    psCommand = psCommand & "if ($avStatus) { "
    psCommand = psCommand & "  Write-Output ('Active Antivirus Engine: ' + $avStatus.displayName + '|COMPLIANT|LOW'); "
    psCommand = psCommand & "} else { "
    psCommand = psCommand & "  Write-Output 'Endpoint Protection: No Active Antivirus Engine Registered!|NON-COMPLIANT|CRITICAL'; "
    psCommand = psCommand & "}; "
    psCommand = psCommand & "$firewall = Get-NetFirewallProfile | Where-Object { $_.Enabled -eq $false }; "
    psCommand = psCommand & "if ($firewall) { "
    psCommand = psCommand & "  Write-Output 'Network Defense: One or more Windows Firewall Profiles Disabled!|NON-COMPLIANT|CRITICAL'; "
    psCommand = psCommand & "} else { "
    psCommand = psCommand & "  Write-Output 'Network Defense: All Windows Firewall Profiles Active|COMPLIANT|LOW'; "
    psCommand = psCommand & "}; "
    
    'Event Log Cleared Tamper Detection
    psCommand = psCommand & "Write-Output '--- FORENSIC AUDIT TRAIL INTEGRITY ---'; "
    psCommand = psCommand & "$logClear = Get-WinEvent -LogName Security | Where-Object { $_.Id -eq 1102 }; "
    psCommand = psCommand & "if ($logClear) { "
    psCommand = psCommand & "  Write-Output ('Security Audit Log: Clear Log Event 1102 Detected (' + $logClear.Count + ' times)!|MALICIOUS ACTIVITY DETECTED|CRITICAL'); "
    psCommand = psCommand & "} else { "
    psCommand = psCommand & "  Write-Output 'Security Audit Log: Event Log Integrity Intact|COMPLIANT|LOW'; "
    psCommand = psCommand & "}; "
    
    ' Close the PowerShell string encapsulation quote
    psCommand = psCommand & """"
    Set wshShell = CreateObject("WScript.Shell")
    Set execProcess = wshShell.Exec(psCommand)
    resultOutput = execProcess.StdOut.ReadAll
    lines = Split(resultOutput, vbCrLf)
    rowNum = 5
    Application.ScreenUpdating = False
    
    'output data stream into the analytics worksheet dashboard
    For Each lineItem In lines
        If Trim(lineItem) <> "" Then
            'Skip structural section separators
            If Left(lineItem, 3) <> "---" Then
                Dim fields() As String
                fields = Split(lineItem, "|")
                
                'Safeguard against array indexing errors on system message anomalies
                If UBound(fields) >= 2 Then
                    currentSheet.Cells(rowNum, 1).Value = fields(0)
                    currentSheet.Cells(rowNum, 2).Value = fields(1)
                    currentSheet.Cells(rowNum, 3).Value = fields(2)
                    
                    'Apply GRC conditional formatting alerting matrices
                    Select Case fields(2)
                        Case "CRITICAL"
                            currentSheet.Range(currentSheet.Cells(rowNum, 1), currentSheet.Cells(rowNum, 3)).Interior.Color = RGB(255, 204, 204)
                            currentSheet.Cells(rowNum, 3).Font.Bold = True
                        Case "MEDIUM"
                            currentSheet.Range(currentSheet.Cells(rowNum, 1), currentSheet.Cells(rowNum, 3)).Interior.Color = RGB(255, 255, 204)
                        Case "LOW"
                            currentSheet.Range(currentSheet.Cells(rowNum, 1), currentSheet.Cells(rowNum, 3)).Interior.Color = RGB(204, 255, 204)
                    End Select
                    rowNum = rowNum + 1
                End If
            End If
        End If
    Next lineItem
    
    currentSheet.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Daily Enterprise Security Operations Automation Complete." & vbCrLf & _
           "System Baselines, Vulnerability Classifications, and Forensics Audited.", _
           vbInformation, "SecOps Orchestration Complete"
End Sub

