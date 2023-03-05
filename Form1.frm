VERSION 5.00
Object = "{6508F09E-5698-11D5-B6D5-0050BA8DB63D}#27.0#0"; "GridAdv.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin AdvGrid.AdvanceGrid AdvanceGrid1 
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3836
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      BeginProperty FontName {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColCount        =   1
      CAP0            =   ""
      ALG0            =   0
      FRM0            =   ""
      WDT0            =   1514,835
      BTN0            =   -1  'True
      CAP1            =   ""
      ALG1            =   0
      FRM1            =   ""
      WDT1            =   1514,835
      BTN1            =   0   'False
      AllowUpdate     =   0   'False
      HeadFontSize    =   8.25
      HeadFontName    =   "MS Sans Serif"
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      EDate           =   44990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send To Grid"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return from Grid"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Click on the column 2 that is last column to activate Add New Row Button to add new row in grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About Me"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset

'*****************************************************
'*****************************************************

'Please Register the OCX (GridAdv.ocx) present in the folder
'called AdvOcx using command Regsvr32 or include in ur project
'from project reference and then browse it
'u need to have ado 2.1 reference in ur project.

'****************************************************
'*****************************************************



Private Sub Command1_Click()
    '*********Fetching recordset from Adv Grid ********
    Dim strreturn As String
    AdvanceGrid1.GridRecordset.MoveFirst
    Do Until AdvanceGrid1.GridRecordset.EOF
      strreturn = strreturn & AdvanceGrid1.GridRecordset.Fields(0).Value & vbCrLf
      AdvanceGrid1.GridRecordset.MoveNext
    Loop
 MsgBox strreturn, , "Values Of First Column"
End Sub

Private Sub Command2_Click()

  '*********Sending recordset to Adv Grid ********

    Set rs = New ADODB.Recordset
    rs.Fields.Append "WithDropDown Grid", adVarChar, 1000
    rs.Fields.Append "WithDropDown Calander", adVarChar, 1000
    rs.Fields.Append "nothing", adVarChar, 1000
    rs.Open

    For i = 1 To 5
         rs.AddNew
         rs.Fields(0).Value = "col 1row " & i
         rs.Fields(1).Value = Date
    Next
    
    AdvanceGrid1.GridRecordset = rs
    
   


End Sub



Private Sub Form_Load()

 '*********Setting properties for dropdown ********
    
  Dim dd1 As ADODB.Recordset
  Set dd1 = New ADODB.Recordset
  


  dd1.Fields.Append "Name", adLongVarChar, 1000
  dd1.Fields.Append "Id", adLongVarChar, 1000

  dd1.Open
  
  dd1.AddNew
  dd1.Fields(0).Value = "first"
  dd1.Fields(1).Value = "f"
  dd1.AddNew
  dd1.Fields(0).Value = "Second"
  dd1.Fields(1).Value = "h"
  dd1.AddNew
  dd1.Fields(0).Value = "0"
  dd1.Fields(1).Value = "j"
  
  

  AdvanceGrid1.DropDownGrid 0, dd1
  
  AdvanceGrid1.DropDownCalander 1
  

'  Command2_Click
    

End Sub



Private Sub mnuabout_Click()
    MsgBox "Hello I'm UMESH DWIVEDI" & vbCrLf _
        & "Please send your feedback to " & vbCrLf _
        & vbCrLf _
        & " umesh909@yahoo.com", vbInformation, "About Me"
End Sub
