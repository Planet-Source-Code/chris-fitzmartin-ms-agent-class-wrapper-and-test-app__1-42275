VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'---------------------------------------------------------
' Singleton
' Wrapper for the MS Agent AgentCtl.DLL.
' Exposes text-to-speech functionality.
'---------------------------------------------------------
' Requires SAPI 4: http://activex.microsoft.com/activex/controls/sapi/spchapi.exe
' Requires install of MS Agent DLL and *.acs" agent file(s).
' http://www.microsoft.com/msagent/downloads.htm
' VB Project needs reference to AgentCtl.DLL. "Microsoft Agent Control 2.0"
' You do *not* need the OCX in the toolbox or on a form !
'---------------------------------------------------------
' Usage example:
'---------------------------------------------------------
'Public Agent As CAgent
'Sub Main()
'Set Agent = New CAgent: Stop
'Agent.Speak "Speak any text. Separate words, with, commas, to, pause. ": Stop
'End Sub
'---------------------------------------------------------

' Top-level MS Agent object. No OCX necessary.
' "Agent" event handlers are in this class.
Private WithEvents objAgent As AgentObjects.Agent
Attribute objAgent.VB_VarHelpID = -1
' main object to program against
Private Character As IAgentCtlCharacterEx

Private Enum eRequestStatus
  eSuccess = 0 'Request successfully completed.
  eFailed = 1 'Request failed.
  ePending = 2 'Request pending (in the queue, but not complete).
  eInterrupted = 3 'Request interrupted.
  eInProgress = 4 'Request in progress.
End Enum

'style bits for character.balloon.style
Private Const BalloonText = 1
Private Const SizeToText = 2
Private Const AutoHide = 4
Private Const AutoPace = 8

Private Const MAKE_INVISIBLE As Integer = -10000    '/ Screen.TwipsPerPixelX  ' character pos. is pixels
Private mlngRestoreXPosition As Long 'pixels
Private msUserName As String
Private mbVisible As Boolean
Private mbBalloonVisible As Boolean
Private mbCharacterLoaded As Boolean
Private msPitch As String
Private miHeight As Integer
Private miWidth As Integer

Public Property Let Pitch(vData As Variant)

  'doc note: The speech engine supplied with Microsoft Agent supports values from 50 to 400.

  ' special tag to set the "Baseline" pitch of the character's voice
  If vData > 200 Then vData = 200 'max
  If vData < 40 Then vData = 40 'min
  
  ' speech engine tag to set pitch at runtime
  msPitch = "\Pit=" & CStr(vData) & "\"
  
End Property

Public Property Get Pitch() As Variant
    Pitch = msPitch
    If Pitch = "" Then Pitch = Character.Pitch
End Property

Public Property Let BalloonVisible(vData As Boolean)
  
If Not mbCharacterLoaded Then Exit Property

  If vData Then
    'Put the spoken text in a ballon popup window.
    Character.Balloon.Style = Character.Balloon.Style Or BalloonText
  Else
    'don't put the spoken text in a ballon popup window
    Character.Balloon.Style = Character.Balloon.Style And (Not BalloonText)
  End If
    
    mbBalloonVisible = vData
End Property

Public Property Get BalloonVisible() As Boolean
    BalloonVisible = mbBalloonVisible
End Property

Public Property Let Visible(vData As Boolean)
  
' there's no direct way to set the character's visibility
' so just move the thing on or off screen.
If Not mbCharacterLoaded Then Exit Property
  
  If vData Then
    Character.Left = mlngRestoreXPosition  'pixels
  Else
    Character.Left = MAKE_INVISIBLE
  End If
  
  mbVisible = vData
    
End Property

Public Property Get Visible() As Boolean
    Visible = mbVisible
End Property

Public Property Let UserName(vData As String)
  'Comma means pause the speech engine for an instant.

  If LCase(vData) = "red" Then vData = "redd pplayer"
  If LCase(vData) = "blue" Then vData = "bloo oo pplayer"
  
  msUserName = vData & ","
End Property

Public Property Get UserName() As String
    UserName = msUserName
End Property

Public Property Get CharacterName() As String
    CharacterName = Character.Name
End Property

Public Sub Speak(sText As String)
    
'You can also include vertical bar characters (|) in the
'Text parameter to designate alternative strings,
' so that the server randomly chooses a different
'string each time it processes the method.
'----------------------------------------------------
'Character.Active property return values
'0 Not the active client.
'1 The active client.
'2 The input-active client. (The topmost character.)
'----------------------------------------------------

If Not mbCharacterLoaded Then Exit Sub

  ' if the chracter is not "ready to speak," then show it.
  ' True = skip the animation prior to "Show"
  If Character.Active <> 2 Then Character.Show (True)
  ' interupt current message with the new message
  Character.Stop
  
  ' say it !
  Character.Speak msPitch & " " & msUserName & " " & sText
    
End Sub

Public Sub LoadCharacter(sCharacter As String)

On Error Resume Next

'' download and install "*.acs" files from:
' http://www.microsoft.com/msagent/downloads.htm#character

' "*.acs" files need to be in the system path or
' put the files in the app.path and load this way.

Dim sCharacterFile As String  ' "*.acs" file
Dim sCharacterName As String
' special object returned by load method
Dim objRequest As IAgentCtlRequest

' comment next line out if acs files are in the system path.
'sCharacterFile = App.Path & "\"

'-----------------------------------------------------------
' These are the free MS characters
' see: http://www.agentry.net
' These are the character files that define a given persona.
'-----------------------------------------------------------
Select Case sCharacter
  Case "Genie"
    sCharacterName = "Genie"
    sCharacterFile = sCharacterFile & "Genie.acs"
  
  Case "Merlin"
    sCharacterName = "Merlin"
    sCharacterFile = sCharacterFile & "Merlin.acs"
  
  Case "Robby"
    sCharacterName = "Robby"
    sCharacterFile = sCharacterFile & "Robby.acs"
  
  Case "Peedy"
    sCharacterName = "Peedy"
    sCharacterFile = sCharacterFile & "Peedy.acs"
  
  Case Else
    ' 2 ways to invoke default character
    'sCharacterName = "default"
    sCharacterName = ""

End Select

Dim bResult As Boolean

' if the full path was specified, check to see if the acs file exists.
If InStr(sCharacterFile, ":") <> 0 Or InStr(sCharacterFile, "\\") <> 0 Then
  bResult = FileExists(sCharacterFile)
  If Not bResult Then
    MsgBox "Speech engine character file was not found: " & vbCrLf & vbCrLf & sCharacterFile
    Exit Sub
  End If
End If

' Unload existing character. Cleans up on a character re-load.
If Not Character Is Nothing Then
  If Character.Name <> "" Then objAgent.Characters.Unload Character.Name
  Set Character = Nothing
End If

'--------------------------------------------------------------------------------------
' There are two ways to load a character. (1) specify name (only) as "default"
' or (2) load characer from an ACS file.
'--------------------------------------------------------------------------------------
' Just load the "default" character.
' The benefit is that you can change the character.
If sCharacterName = "" Or sCharacterName = "default" Then
   objAgent.Characters.Load ("default")
   Set Character = objAgent.Characters("default")
   If Character Is Nothing Then Exit Sub
Else
  ' create a character from the character file "sCharacterFile"
  Set objRequest = objAgent.Characters.Load(sCharacterName, sCharacterFile)
  'validate character loaded
  If objRequest.Status <> eRequestStatus.eSuccess Then Exit Sub
  Set Character = objAgent.Characters(sCharacterName)
End If
'-----------------------------------------------------------

'-----------------------------------------------------------
' english
'-----------------------------------------------------------
'Set the character's language ID (by setting the character's LanguageID before
'using the Speak method to ensure appropriate text display within the word balloon.
Character.LanguageID = &H409  '= 1033
If Not Character Is Nothing Then mbCharacterLoaded = True
  Debug.Print TypeName(Character)  ' "IAgentCtlCharacterEx"

End Sub

Private Sub SetupCharacterDefaults()

If Not mbCharacterLoaded Then Exit Sub

'-------------------------------------------------------------
' position and size
'-------------------------------------------------------------
' character starts out here
MoveTopRight
' redundant, but theoretically necessary
mlngRestoreXPosition = Character.Left
' make the character smaller than the default
Character.Height = Character.Height * 0.75
Character.Width = Character.Width * 0.75
'----------------------------------------------------------------------------
' default visibility: character and balloon
'----------------------------------------------------------------------------
'Me.Visible = False
'Me.BalloonVisible = False
'Me.ShowBalloonTextOnly
'----------------------------------------------------------------------------

'-------------------------------------------------------------
' Balloon style & behavior
'-------------------------------------------------------------
With Character.Balloon
  ' Scroll in the balloon as the text is spoken
  .Style = .Style Or AutoPace
  ' jumps balloon text to the end. no scrolling as you go.
  '.Style = .Style And (Not AutoPace)
  
  ' make the ballon size proportional to message. balloon resizes dramatically.
  '.Style = .Style Or SizeToText
  ' scroll message in small balloon. don't resize balloon.
  .Style = .Style And (Not SizeToText)
  
  ' Hide the balloon after every ".Speak"
  .Style = .Style Or AutoHide
  ' leave the balloon visible after speech
  '.Style = .Style And (Not AutoHide)
End With
'-------------------------------------------------------------
' Sound
'-------------------------------------------------------------
' This property determines whether sound effects included as a
' part of a character's animations will play when an animation plays.
Character.SoundEffectsOn = False  'or True
'-------------------------------------------------------------

'-------------------------------------------------------------
' Voice -- I like the range from 110 to 140 Hz for "clarity"
'-------------------------------------------------------------
Me.Pitch = 130  ' range = 40 to 200 Hz

' command to make the character ready to speak.
Character.Show (True) 'show and skip animation prior to .Show

End Sub

Private Sub Class_Terminate()
Set Character = Nothing
Set objAgent = Nothing
End Sub

Public Sub Test()

On Error Resume Next

Character.Stop
Character.SoundEffectsOn = True
Character.Speak "Hello, " & Me.UserName & ". My name is " & Character.Name & ", , " & Character.Description '& ". Dave, you know, how much enthusiasm, I have, for the, mission. But, I am sorry, Dave,. I can't open, the pod bay door. ha ha ha, ha ha ha, ha ha ha "

' test the ability to make the character gesture and such

Dim i As Integer
Dim AnimationName As Variant
Dim Names As IAgentCtlAnimationNames
Set Names = Character.AnimationNames
  Debug.Print TypeName(Names)   ' IAgentCtlAnimationNames

' shows what this agent can do
For Each AnimationName In Names
  Animate CStr(AnimationName)
Next

Animate "RestPose"

If Err <> 0 Then Beep

End Sub

Public Sub ShowPropertyTab()
' shows built-in property sheet tabs
  objAgent.PropertySheet.Visible = True
End Sub

Public Sub Animate(sActionName As String)
'Character.Stop
Character.Play (sActionName)
End Sub

Private Sub AddCustomMenus()

'-----------------------------------------------------
' See "objAgent_Command" event handler
'-----------------------------------------------------

' cleanup old menus
'Character.RemoveAll

If Not mbCharacterLoaded Then Exit Sub

With Character.Commands
  .Add "UserMenuTest", "&Test Me"
  .Add "UserMenuAdvancedOptions", "&Advanced Options"
  .Add "UserMenuChooseAgent", "&Choose Agent"
  .Add "UserMenuShowSpeechInputCommandWindow", "Input Window"
  .Add "UserMenuDecreasePitch", "Voice Pitch &DOWN"
  .Add "UserMenuIncreasePitch", "Voice Pitch &UP"
  .Add "UserMenuHideBaloon", "H&ide Balloon"
  .Add "UserMenuShowBalloon", "&Show Balloon"
  .Add "UserMenuTurnOnSound", "T&urn Sound ON"
  .Add "UserMenuTurnOffSound", "Tu&rn Sound OFF"
  .Add "UserMenuShowBalloonTextOnly", "Show Balloon Text Only"
  .Add "UserMenuShowBalloonAndAgent", "Show Balloon And Agent"
  .Add "UserMenuOriginalSize", "Full Size"
  .Add "UserMenu3_4Size", Chr(190) & " Size"
  .Add "UserMenuSaySomething", "Say Something !"
  '.Add "UserMenuX", "X"
End With

End Sub

Private Sub objAgent_Command(ByVal UserInput As Object)

'--------------------------------------------------
' Custom event handler for user defined menu item.
'--------------------------------------------------

Select Case UserInput.Name
  Case "UserMenuTest"
    Me.Test
  Case "UserMenuAdvancedOptions"
    Me.ShowPropertyTab
  Case "UserMenuChooseAgent"
    Me.ChooseCharacter
  Case "UserMenuDecreasePitch"
    Me.Pitch = Replace(Replace(Me.Pitch, "\Pit=", ""), "\", "") - 10
  Case "UserMenuIncreasePitch"
    Me.Pitch = Replace(Replace(Me.Pitch, "\Pit=", ""), "\", "") + 10
  Case "UserMenuHideBaloon"
    Me.BalloonVisible = False
  Case "UserMenuShowBalloon"
    Me.BalloonVisible = True
  Case "UserMenuTurnOnSound"
    Character.SoundEffectsOn = True
  Case "UserMenuTurnOffSound"
    Character.SoundEffectsOn = False
  Case "UserMenuShowBalloonTextOnly"
    Me.ShowBalloonTextOnly
  Case "UserMenuShowBalloonAndAgent"
    Me.ShowBalloonTextAndCharacter
  Case "UserMenuOriginalSize"
    Character.Height = Character.OriginalHeight
    Character.Width = Character.OriginalWidth
  Case "UserMenu3_4Size"
    Character.Height = Character.OriginalHeight * 0.75
    Character.Width = Character.OriginalWidth * 0.75
  Case "UserMenuSaySomething"
    Dim sResult As String
    sResult = InputBox("Type Something", "Test Text-to-Speech Agent", "Hello World")
    If sResult <> "" Then Me.Speak sResult
  Case "UserMenuShowSpeechInputCommandWindow"
    Me.ShowSpeechInputWindow
  Case Else
    Debug.Print UserInput.Name & "  " & "Sub objAgent_Command ? ? ?"
End Select

  Debug.Print TypeName(UserInput) ' type of late bound object

End Sub

Private Function FileExists(Filename As String) As Boolean
On Error Resume Next

Dim sResult As String

sResult = Dir(Filename)
DoEvents: DoEvents
If sResult <> "" Then
    FileExists = True
Else
    FileExists = False
End If

End Function

Private Sub objAgent_DefaultCharacterChange(ByVal GUID As String)
'-------------------------------------------------------------------------
' For this to work, you have to load the "default" character on startup.
'-------------------------------------------------------------------------
 Debug.Print "Character Changed -- " & Character.Name
End Sub

'Description
' Sets the baseline average talking speed of the speech output.
' The speech engine supplied with Microsoft Agent supports values from 50 to 250.
'Syntax
'  \Spd=number\
'  \Spd=218\
'  number Baseline average talking speed, in words per minute.
'
'Description
'  Sets the baseline speaking volume of the speech output.
'Syntax
' \Vol=number\
' \Vol=65530\
'  number Baseline speaking volume: 0 is silence and 65535 is maximum volume.

Private Sub objAgent_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
  ' event handler fires when user chooses "Hide" from the popup menu.
  ' if the user "Hides" using the popup menu, keep hidden
  'Me.Visible = False
End Sub

Public Sub ChooseCharacter()
'-------------------------------------------------------------------------
' For this to work, you have to load the "default" character on startup.
'-------------------------------------------------------------------------
' shows built-in dialog to change characters
  If Not objAgent Is Nothing Then
    objAgent.ShowDefaultCharacterProperties
  End If
End Sub

Private Sub Class_Initialize()

Debug.Print TypeName(Me); " Object Initialize"

'-----------------------------------------------------------------------------------------------------
' create a top-level Agent object
'-----------------------------------------------------------------------------------------------------
Set objAgent = New AgentObjects.Agent
  Debug.Print TypeName(objAgent)  'Agent

' From online doc: You also need to use the Connected property before you reference any
' Agent methods or properties. If this is not done, MS Agent will raise an error indicating
' that it was unable to start (error code 80042502).
If Not objAgent.Connected Then
  Debug.Print "Agent object was not created in class init."
  Exit Sub
End If
'-----------------------------------------------------------------------------------------------------

'--------------------------------------------------------
' Load character data from file or via default mechanism
'--------------------------------------------------------
LoadCharacter "default" 'default character
'LoadCharacter ""        'default character
'LoadCharacter "Merlin"  'med-evil magician
'LoadCharacter "Genie"   'arab genie
'LoadCharacter "Robby"   'silver robot
'LoadCharacter "Peedy"   'green bird
'--------------------------------------------------------

If Not mbCharacterLoaded Then
  Debug.Print "A Character was not loaded."
  Exit Sub
End If

' my preferences
SetupCharacterDefaults

Me.Speak "Speech Engine Initialized"

AddCustomMenus

miHeight = Character.Height
miWidth = Character.Width

End Sub

Public Sub Think(sText As String)
If Not mbCharacterLoaded Then Exit Sub
  If sText = "" Then Exit Sub
  Character.Think sText
End Sub

Public Sub ShowBalloonTextOnly()
If Not mbCharacterLoaded Then Exit Sub
Character.Height = 1
Character.Width = 1
End Sub

Public Sub ShowBalloonTextAndCharacter()
If Not mbCharacterLoaded Then Exit Sub
Character.Height = miHeight
Character.Width = miWidth
End Sub

Public Sub MoveCharacter(pixX As Integer, pixY As Integer)

'Speed Optional.
'A Long integer value specifying in milliseconds how quickly the character's frame moves.
'The default value is 1000.
'Specifying zero (0) moves the frame without playing an animation.

If Not mbCharacterLoaded Then Exit Sub

Character.Stop
Character.MoveTo pixX, pixY ', 1000

'save
mlngRestoreXPosition = pixX

End Sub
Public Sub MoveTopLeft()
MoveCharacter 0, 0
End Sub
Public Sub MoveTopRight()
MoveCharacter (Screen.Width / Screen.TwipsPerPixelX) - 120, 0
End Sub
Public Sub MoveBottomLeft()
MoveCharacter 0, (Screen.Height / Screen.TwipsPerPixelY) - 120
End Sub
Public Sub MoveBottomRight()
MoveCharacter (Screen.Width / Screen.TwipsPerPixelX) - 120, (Screen.Height / Screen.TwipsPerPixelY) - 200
End Sub
Public Sub MoveCenter()
MoveCharacter ((Screen.Width / Screen.TwipsPerPixelX) / 2) - 60, ((Screen.Height / Screen.TwipsPerPixelY) / 2) - 50
End Sub

Public Sub DumpAnimationList()

Dim AnimationName As Variant
Dim Names As IAgentCtlAnimationNames
Dim iCount As Integer
Set Names = Character.AnimationNames
  Debug.Print TypeName(Names) 'IAgentCtlAnimationNames

If Not mbCharacterLoaded Then Exit Sub

Debug.Print "---------------------------------------"
Debug.Print Character.Name & ": animation name list . . . "
Debug.Print "---------------------------------------"

For Each AnimationName In Names
  Debug.Print AnimationName
  ' do it
  'Me.Animate AnimationName
  iCount = iCount + 1
Next

Debug.Print "---------------------------------------"
Debug.Print "There are " & iCount & " Animations"
Debug.Print "---------------------------------------"

End Sub

Public Sub ShowSpeechInputWindow()
objAgent.CommandsWindow.Visible = True
End Sub
