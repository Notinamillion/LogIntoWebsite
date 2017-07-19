set w=wscript.CreateObject("wscript.shell")
w.run "chrome.exe -incognito"
wscript.sleep (3700)
w.SendKeys "login.microsoftonline.com/"
w.SendKeys "{ENTER}"
wscript.sleep (4500)
w.SendKeys "test@DOMAIN.com"
w.SendKeys "{TAB}"
wscript.sleep (3500)
w.SendKeys "PASSWORD"
w.SendKeys "{ENTER}"

wscript.sleep (11500)

w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"

w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"
w.SendKeys "{TAB}"