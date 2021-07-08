' Send Password

Set WshShell = WScript.CreateObject("WScript.Shell")

Mdp = "?VHHPF"&"""{"""&"48KBIEs"&"}"&"-"

wscript.sleep 2500


For Counter = 1 To Len (Mdp)
    'do something to each character in string
    'here we'll msgbox each character
    Letter =  Mid(Mdp , Counter, 1)
	WshShell.SendKeys Letter
Next
 

