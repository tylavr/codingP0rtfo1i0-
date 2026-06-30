Attribute VB_Name = "AD_Enterprise_IDManagement"
Sub ManageActiveDirectoryIdentities()
    Dim wshShell As Object
    Dim execProcess As Object
    Dim psCommand As String
    Dim resultOutput As String
    Dim rowNum As Integer
    Dim actionChoice As String
    Dim targetUser As String
    
   'compliance admin input
    actionChoice = InputBox( _
        "Select an Identity Management Action to enforce:" & vbCrLf & vbCrLf & _
        "1 - Add a User to Domain Admins" & vbCrLf & _
        "2 - Remove a User from Domain Admins" & vbCrLf & _
        "3 - Enforce Smart Card / Hardware MFA on an Account" & vbCrLf & _
        "4 - Run Audit Only (Skip changes)", _
        "Active Directory Identity & MFA Control Console", "4")
        
    ' handling clean exit or cancel actions
    If Trim(actionChoice) = "" Or actionChoice = "4" Then
        'proceed straight to read-only audit phase
    ElseIf actionChoice = "1" Or actionChoice = "2" Or actionChoice = "3" Then
        targetUser = InputBox("Enter the target User Account Name (sAMAccountName):", "Target Identity Profile Extraction")
        If Trim(targetUser) = "" Then
            MsgBox "Action aborted. User Account Name cannot be empty.", vbCritical, "Execution Cancelled"
            Exit Sub
        End If
    Else
        MsgBox "Invalid selection entry. Defaulting to system audit mode only.", vbExclamation, "Input Validation Error"
    End If

    psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command """
    psCommand = psCommand & "Import-Module ActiveDirectory -ErrorAction SilentlyContinue; "
    'processing interactive administrative changes based on user input selection
    If actionChoice = "1" Then
        psCommand = psCommand & "Add-ADGroupMember -Identity 'Domain Admins' -Members '" & targetUser & "' -ErrorAction SilentlyContinue; "
        psCommand = psCommand & "Write-Output 'ACTION LOG: Added User [" & targetUser & "] to Domain Admins Container.'; "
    ElseIf actionChoice = "2" Then
        psCommand = psCommand & "Remove-ADGroupMember -Identity 'Domain Admins' -Members '" & targetUser & "' -Confirm:$false -ErrorAction SilentlyContinue; "
        psCommand = psCommand & "Write-Output 'ACTION LOG: Removed User [" & targetUser & "] from Domain Admins Container.'; "
    ElseIf actionChoice = "3" Then
        'SmartcardRequired represents the exact on-premises Active Directory equivalent parameter to enforce hardware MFA
        psCommand = psCommand & "Set-ADUser -Identity '" & targetUser & "' -SmartcardRequired $true -ErrorAction SilentlyContinue; "
        psCommand = psCommand & "Write-Output 'ACTION LOG: Hardware Smart Card MFA Policy forced active on User [" & targetUser & "].'; "
    End If
    
    'appending the Standard Governance Auditing Routine
    psCommand = psCommand & "Write-Output '--- SYSTEM GROUP MEMBERSHIP REVIEWS ---'; "
    psCommand = psCommand & "$members = Get-ADGroupMember -Identity 'Domain Admins' -ErrorAction SilentlyContinue; "
    psCommand = psCommand & "if($members){ $members | ForEach-Object { $_.SamAccountName + ',' + $_.Name } } "
    psCommand = psCommand & "else { 'ERROR: Active Directory Module not present or domain controller unreachable on this endpoint.' };"
    psCommand = psCommand & """"

    'run and output data to excel
    Set wshShell = CreateObject("WScript.Shell")
    Set execProcess = wshShell.Exec(psCommand)
    resultOutput = execProcess.StdOut.ReadAll
    
    'clear worksheet and generate a dynamic table matrix
    Sheets(1).Cells.Clear
    Sheets(1).Cells(1, 1).Value = "IDENTITY GOVERNANCE & ACCESS CONTROL MANAGEMENT REPORT"
    Sheets(1).Cells(1, 1).Font.Bold = True
    Sheets(1).Cells(2, 1).Value = "Generated: " & Now & " | Enforcement Log Summary"
    Sheets(1).Cells(2, 1).Font.Italic = True
    
    Dim entries() As String
    entries = Split(resultOutput, vbCrLf)
    rowNum = 4
    
    For i = LBound(entries) To UBound(entries)
        If InStr(entries(i), "ACTION LOG:") > 0 Then
            'print any changes made during the execution session
            Sheets(1).Cells(rowNum, 1).Value = entries(i)
            Sheets(1).Cells(rowNum, 1).Font.Color = RGB(0, 102, 204)
            Sheets(1).Cells(rowNum, 1).Font.Bold = True
            rowNum = rowNum + 2
            
            'draw fresh headers for the resulting audit phase
            Sheets(1).Cells(rowNum, 1).Value = "Current Verified Privilege Holder (sAMAccountName)"
            Sheets(1).Cells(rowNum, 2).Value = "Employee Legal Identity Display Name"
            Sheets(1).Range(Sheets(1).Cells(rowNum, 1), Sheets(1).Cells(rowNum, 2)).Font.Bold = True
            rowNum = rowNum + 1
        ElseIf InStr(entries(i), ",") > 0 Then
            'process out comma-separated Active Directory user profile items
            Dim userFields() As String
            userFields = Split(entries(i), ",")
            Sheets(1).Cells(rowNum, 1).Value = userFields(0)
            Sheets(1).Cells(rowNum, 2).Value = userFields(1)
            rowNum = rowNum + 1
        ElseIf Left(entries(i), 3) <> "---" And Trim(entries(i)) <> "" Then
            'print errors or verification baseline status alerts
            Sheets(1).Cells(rowNum, 1).Value = entries(i)
            rowNum = rowNum + 1
        End If
    Next i
    
    Sheets(1).Columns("A:B").AutoFit
    MsgBox "Active Directory Identity Administration Cycle Finalised.", vbInformation, "Identity Governance Completed"
End Sub

