VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL String Converter"
   ClientHeight    =   3312
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5184
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3312
   ScaleWidth      =   5184
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Convert and Copy to Clipboard"
      Height          =   300
      Left            =   1116
      TabIndex        =   1
      Top             =   2904
      Width           =   2892
   End
   Begin VB.TextBox Text1 
      Height          =   2652
      Left            =   96
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5004
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim strSQL As String
    Dim fs As Object, a As Object
    Dim icnt As Long
    strSQL = Text1.Text
    strSQL = Replace(strSQL, """", """""")
    strSQL = Replace(strSQL, "      ", " ")
    strSQL = Replace(strSQL, "     ", " ")
    strSQL = Replace(strSQL, "    ", " ")
    strSQL = Replace(strSQL, "   ", " ")
    strSQL = Replace(strSQL, "  ", " ")
    strSQL = Replace(strSQL, " ", " ")
    Open "c:\temp.tmp" For Output As #1
    Print #1, strSQL
    Close #1
    strSQL = """"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OpenTextFile("c:\temp.tmp")
    Do While Not a.AtEndOfStream
        icnt = icnt + 1
        If icnt < 5 Then
            strSQL = strSQL & " " & a.readline
        Else
            strSQL = strSQL & """ & _ " & vbCrLf & """" & a.readline
            icnt = 0
        End If
    Loop
    a.Close
    Kill "c:\temp.tmp"
    Clipboard.SetText (strSQL)
    MsgBox "The SQL string is now in your clipboard. It is now ready to be Pasted as a Visual Basic string.", vbOKOnly, "Complete"
    End
End Sub

