Attribute VB_Name = "modGeneral"
Public Const SettingsPath = "Software\MarisBrowser\Settings"
Public sam As String
Option Explicit
Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Global Const GWW_HWNDPARENT = (-8)



Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop

End Sub

Public Sub options()
With Form3
.Check2.Value = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Speed Bar", "1")
.Check1.Value = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Show Pop", "1")
.Check3.Value = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "On Startup", "1")
.Text1.Text = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "info", "Click Browse Button")
.Text2.Text = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "url", "www.geocities.com/Mariskan")
End With
End Sub


