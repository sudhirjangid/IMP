Dim App
Set App = CreateObject("QuickTest.Application")

App.Launch
App.Visible = True

App.Open "C:\Users\sujangid\Desktop\178055(SudhirJangid)\OpenCart", True
Set qtTest = App.Test
qtTest.Settings.Run.OnError = "NextStep"
qtTest.Run
App.Test.Settings.Run.ObjectSyncTimeOut = 30000


ways to call an action things we can do on an action --exit also
different extensions 


scripting - functions

recovery scenario-- in between close the application -- handle object not found 

