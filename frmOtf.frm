VERSION 5.00
Begin VB.Form frmOtf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO Recordset on the fly"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Create the recordset"
      Default         =   -1  'True
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmOtf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim oRs As New ADODB.Recordset
Dim Otf As New OnTheFlyADO_Recordset
Dim i As Integer
Dim str As String

'Define the structure of the recordset - remember that all fields will be of the same type : VarChar
Otf.DefineRecordsetFields "Name##20;Phone Number##10;Country##20"

'Let's add some records ...
Otf.AddRecord "Marcos;89897676;Brazil"
Otf.AddRecord "Johnny;90983434;Brazil"
Otf.AddRecord "Anderson;89892030;USA"

'Set it to a new recordset (no necessary, but if you need to do it...)
Set oRs = Otf.Recordset

'Show the records
oRs.MoveFirst
While Not oRs.EOF
    str = str & oRs("Name") & "  -  " & oRs("Phone Number") & "  -  " & oRs("Country") & vbCrLf
    oRs.MoveNext
Wend

MsgBox str

Set oRs = Nothing
Set Otf = Nothing
End Sub
