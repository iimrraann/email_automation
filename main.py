import subprocess
import pandas as pd


addresses = ['']

for address in addresses:
    email_to = address
    email_cc = address
    email_subject = 'subject'
    email_body = 'body'

    command = f'''
    $outlook = new-object -comobject outlook.application
    $email = $outlook.CreateItem(0)
    $email.To = "{email_to}"
    $email.CC = "{email_cc}"
    $email.Subject = "{email_subject}"
    $email.Body = "{email_body}" 

    Write-Host "Email Created"
    $email.Send()
    $outlook.Quit()
    Write-Host "Email Sent"
    Start-Sleep -Seconds 10
        '''
    
    process = subprocess.Popen(['powershell', '-Command', command], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    stdout, stderr = process.communicate()

    # Decode the output and errors (they come as bytes)
    print("Output:\n", stdout.decode())
    print("Error:\n", stderr.decode())

    break
