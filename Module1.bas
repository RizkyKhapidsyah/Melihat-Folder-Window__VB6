Attribute VB_Name = "Module1"
Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
lpOperation As String, ByVal lpFile As String, ByVal _
lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1


