Attribute VB_Name = "basGlobal"
Option Explicit

' singleton class
Public Agent As CAgent
Sub Main()
Set Agent = New CAgent
Agent.UserName = "blue"
Form1.Show
End Sub

