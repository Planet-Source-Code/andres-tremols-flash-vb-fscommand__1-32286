VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlashControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The purpose of this class is to use callbyname to get
'the flash movie's FSCOMMAND without using Select Case or
'IF's. Each procedure here has the same name of the Flash
'FSCOMMAND command parameter, if the procedure needs
'arguments those are passed through the arguments
'parameter of fscommand









Public Sub BUTTON1(args As String)
'note that the vb procedures need at least an argument
'although you wont use it, otherwise an error will occur.

MsgBox "hello world"
End Sub

Public Sub BUTTON2(args As String)
MsgBox "Info coming from Flash textbox: " & args
End Sub

Public Sub BUTTON1RollOver(args As String)
FlashVb.Label1.Caption = "Rolling Over Flash Button1"

End Sub

Public Sub BUTTON2RollOver(args As String)
FlashVb.Label1.Caption = "Type something in the text field and click"

End Sub
