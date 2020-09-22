VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form FlashVb 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Vb Textbox sending Data to Flash Movie (Type Anything below)"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   4800
      Width           =   2895
      Begin VB.Label Label1 
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Flash Mouse X:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Flash Mouse Y:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5640
      Top             =   5520
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash1 
      Height          =   2655
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _cx             =   7223
      _cy             =   4683
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
      Stacking        =   "below"
   End
End
Attribute VB_Name = "FlashVb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private FlashCtrl As New FlashControl




Private Sub flash1_FSCommand(ByVal command As String, ByVal args As String)
'this maps the fscommand to the methods of the FlashControl  Class

CallByName FlashCtrl, command, VbMethod, args

End Sub

Private Sub Form_Load()
flash1.Movie = App.Path & "\buttons2.swf"
flash1.Play

End Sub

 
Private Sub Text1_Change()
'Demo of the setVariable method, it can set flash movie variables, in this case
'its a textfield variable but its not limited to this

flash1.SetVariable "textfieldX", Text1.Text
End Sub

Private Sub Timer1_Timer()
'The timer demonstrates the GetVariable method, the mouse coordinates property is
'converted to variable in the flash movie, the VB timer keeps querying those variables.

Label1.Caption = flash1.GetVariable("x")
Label6.Caption = flash1.GetVariable("y")

End Sub
