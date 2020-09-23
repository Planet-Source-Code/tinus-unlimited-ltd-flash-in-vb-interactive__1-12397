VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Interact with Flash"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmeOpt 
      Caption         =   "Type of Flash:"
      Height          =   1035
      Left            =   4920
      TabIndex        =   2
      Top             =   60
      Width           =   3795
      Begin VB.OptionButton optFlash 
         Caption         =   "&interact with an animation in Flash"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optFlash 
         Caption         =   "&a form in Flash"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   2595
      End
   End
   Begin VB.Frame frmeFlash 
      Caption         =   "Flash:"
      Height          =   6735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4515
         _cx             =   4464412
         _cy             =   4467693
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin VB.Frame frmeFORM 
      Caption         =   "Form Settings:"
      Height          =   5595
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   3795
      Begin VB.CommandButton cmdGetMessage 
         Caption         =   "&Get"
         Height          =   315
         Left            =   3240
         TabIndex        =   15
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton cmdGetSubject 
         Caption         =   "&Get"
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   780
         Width           =   495
      End
      Begin VB.CommandButton cmdGetEmail 
         Caption         =   "&Get"
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdSetMessage 
         Caption         =   "&Set"
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton cmdSetSubject 
         Caption         =   "&Set"
         Height          =   315
         Left            =   2700
         TabIndex        =   11
         Top             =   780
         Width           =   495
      End
      Begin VB.CommandButton cmdSetEmail 
         Caption         =   "&Set"
         Height          =   315
         Left            =   2700
         TabIndex        =   10
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox txtMessage 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1260
         Width           =   2475
      End
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "subject"
         Top             =   780
         Width           =   2475
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "e-mail"
         Top             =   300
         Width           =   2475
      End
      Begin VB.Label lblBtnSub 
         Caption         =   "btnSub is: down"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblBtnClear 
         Caption         =   "btnClear is: down"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   2940
         Width           =   2415
      End
   End
   Begin VB.Frame frmeANIMATION 
      Caption         =   "Animation Settings:"
      Height          =   5595
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   3795
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Width           =   915
      End
      Begin VB.HScrollBar hscrollEff3 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   360
         TabIndex        =   31
         Top             =   3180
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollEff2 
         Height          =   255
         Left            =   120
         Max             =   1
         TabIndex        =   29
         Top             =   2820
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollEff1 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   100
         TabIndex        =   27
         Top             =   2460
         Value           =   100
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollYScale 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   25
         Top             =   1680
         Value           =   100
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollXScale 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   23
         Top             =   1320
         Value           =   100
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollY 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   365
         Min             =   87
         TabIndex        =   21
         Top             =   960
         Value           =   87
         Width           =   1695
      End
      Begin VB.HScrollBar hscrollX 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   266
         Min             =   33
         TabIndex        =   19
         Top             =   600
         Value           =   33
         Width           =   1695
      End
      Begin VB.Label lblNOTALL 
         Caption         =   "[ NOT ALL EFFECTS ARE CODED! ]"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label lblEff3 
         Caption         =   "Rotation: 0"
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label lblEff2 
         Caption         =   "Visibility: 1"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblEff1 
         Caption         =   "Alpha: 100"
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label lblYScale 
         Caption         =   "Y-scale: 100"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label lblXScale 
         Caption         =   "X-scale: 100"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label lblY 
         Caption         =   "Y: 87"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblX 
         Caption         =   "X: 33"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lblBall 
         Caption         =   "Ball [movie clip in Flash]:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frameNUM1 As Long
Dim frameNUM2 As Long
Dim Total As String

Dim X As Long
Dim Y As Long
Dim XScale As Long
Dim YScale As Long

'// OKAY, HERE's SOME INTERACTION WITH FLASH
'// I CODED A FEW EFFECT AND A TEXT INTERACTION
'// BUT NOT ALL EFFECT ARE CODED HERE, SO FIND OUT
'// FOR YOURSELF WHAT FLASH CAN DO FOR YOU...
'// ANIMATION, GUI, FORM, ASP, IT'S ALL POSSIBILE
'// HAVE FUN WITH IT, TINUS.
'// GO AHEAD AND STEAL MY CODE !
'// (C) TINUS UNLIMITED LTD. vliertstraat@wish.net

Private Sub Form_Load()
 'Load the Flash in the control
 flash.Movie = App.Path & "\flash.swf"
 'Setting the frame numbers so we can switch inside Flash
 frameNUM1 = 1
 frameNUM2 = 16
 'set Settings-frame in VB
 frmeFORM.Visible = True
 frmeANIMATION.Visible = False
End Sub

Private Sub optFlash_Click(Index As Integer)
 If Index = 1 Then
  'set Flash frame to ANIMATION
  flash.GotoFrame frameNUM2
  'set Settings to ANIMATION
  frmeANIMATION.Visible = True
  frmeFORM.Visible = False
  'Reset Postion, Scale and Effects of Ball [Movie clip in Flash]
  cmdReset.Value = True
 Else
   'set Flash frame to FORM
  flash.GotoFrame frameNUM1
  frmeFORM.Visible = True
  frmeANIMATION.Visible = False
 End If
End Sub

'*********** Form Flash *****************

'The Flash triggered a FS COMMAND
Private Sub flash_FSCommand(ByVal command As String, ByVal args As String)
 'I put two button in the form btnSub and btnClear
 'The actions i connected to the buttons (in the Flash) are:
 'On (PRESS) -> FS Command "down"
 'On (RELEASE) -> FS Command "up"
 'That's all, the rest is handled by VB:
 If command = "btnSub" Then
  If args = "up" Then
   lblBtnSub.Caption = "btnSub is: pressed"
   'Just wait to be released
  Else
   lblBtnSub.Caption = "btnSub is: down"
   'gather the information and send an e-mail
   cmdGetEmail.Value = True                 'Click button
   cmdGetSubject.Value = True               'Click button
   cmdGetMessage.Value = True               'Click button
   Total = "start mailto:" & Trim(txtEmail.Text)
   Total = Total & "?subject=" & Trim(txtSubject.Text)
   Total = Total & "?body=" & Trim(txtMessage.Text)
   dummy = Shell(Total, vbNormalFocus)      'Shell out to default e-mail app.
  End If
 Else
  If args = "up" Then
   lblBtnClear.Caption = "btnClear is: pressed"
   'clear all textboxes in VB
   txtEmail.Text = "": txtSubject.Text = "": txtMessage.Text = ""
   'transfer empty text from VB to Flash
   cmdSetEmail.Value = True                 'Click button
   cmdSetSubject.Value = True               'Click button
   cmdSetMessage.Value = True               'Click button
  Else
   lblBtnClear.Caption = "btnClear is: down"
   'nothing to do...
  End If
 End If
End Sub

'GET all Text from FORM
Private Sub cmdGetEmail_Click()
 'In the Flash the're 3 textboxes, I gave them a name,
 'the GetVariable get's the value of the textbox
 txtEmail.Text = flash.GetVariable("txtEmail")
End Sub
Private Sub cmdGetSubject_Click()
 txtSubject.Text = flash.GetVariable("txtSubject")
End Sub
Private Sub cmdGetMessage_Click()
 txtMessage.Text = flash.GetVariable("txtMessage")
End Sub

'SET all Text from FORM
Private Sub cmdSetEmail_Click()
 flash.SetVariable "txtEmail", txtEmail.Text
End Sub
Private Sub cmdSetSubject_Click()
 flash.SetVariable "txtSubject", txtSubject.Text
End Sub
Private Sub cmdSetMessage_Click()
 flash.SetVariable "txtMessage", txtMessage.Text
End Sub

'*********** End of Form in Flash *****************

'*********** Animation in Flash *****************

Private Sub hscrollX_Change()
 lblX.Caption = "X: " & hscrollX.Value
 'target: movie clip called Ball
 'property: X_postition = 0
 flash.TSetProperty "Ball", 0, Str(hscrollX.Value)
End Sub

Private Sub hscrollY_Change()
 lblY.Caption = "Y: " & hscrollY.Value
 'property: Y_postition = 1
 flash.TSetProperty "Ball", 1, Str(hscrollY.Value)
End Sub

Private Sub hscrollXScale_Change()
 lblXScale.Caption = "X-Scale: " & hscrollXScale.Value
 'property: X_Scale = 2
 flash.TSetProperty "Ball", 2, Str(hscrollXScale.Value)
End Sub

Private Sub hscrollYScale_Change()
 lblYScale.Caption = "Y-Scale: " & hscrollYScale.Value
 'property: Y_Scale = 3
 flash.TSetProperty "Ball", 3, Str(hscrollYScale.Value)
End Sub

Private Sub hscrollEff1_Change()
 lblEff1.Caption = "Alpha: " & hscrollEff1.Value
 'property: Alpha = 6
 flash.TSetProperty "Ball", 6, Str(hscrollEff1.Value)
End Sub

Private Sub hscrollEff2_Change()
lblEff2.Caption = "Visibility: " & hscrollEff2.Value
'property: Visibility = 7
 flash.TSetProperty "Ball", 7, Str(hscrollEff2.Value)
End Sub

Private Sub hscrollEff3_Change()
lblEff3.Caption = "Rotation: " & hscrollEff3.Value & Chr(186)
'property: Rotation = 10
 flash.TSetProperty "Ball", 10, Str(hscrollEff3.Value)
End Sub

Private Sub cmdReset_Click()
 'Reset values
 hscrollX.Value = hscrollX.Min + (hscrollX.Max - hscrollX.Min) / 2
 hscrollY.Value = hscrollY.Min + (hscrollY.Max - hscrollY.Min) / 2
 hscrollXScale.Value = hscrollXScale.Min + (hscrollXScale.Max - hscrollXScale.Min) / 2
 hscrollYScale.Value = hscrollYScale.Min + (hscrollYScale.Max - hscrollYScale.Min) / 2
 hscrollEff1.Value = hscrollEff1.Max
 hscrollEff2.Value = hscrollEff2.Max
 hscrollEff3.Value = hscrollEff3.Min
End Sub

'*********** End of Animation in Flash *****************
