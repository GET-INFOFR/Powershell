Start-Sleep -Milliseconds 2000

[System.Windows.Forms.SendKeys]::SendWait("cm-cic\lesiresy2")
Start-Sleep -Milliseconds 750
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
Start-Sleep -Milliseconds 20
[System.Windows.Forms.SendKeys]::SendWait("MOnMot2P@sse:072021!!")
Start-Sleep -Milliseconds 750
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 750

