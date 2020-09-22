VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucReportList test"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Change row"
      Height          =   510
      Left            =   6270
      TabIndex        =   1
      Top             =   180
      Width           =   1560
   End
   Begin Test.ucReportList ucReportList1 
      Height          =   4935
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   8705
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim lItm As Long
  Dim lCol As Long
  Dim sRow As String
    
    ucReportList1.AddHeader 100, LeftJustify, "Column 1"  ' First column always <LeftJustify> (*)
    ucReportList1.AddHeader 60, RightJustify, "Column 2"
    ucReportList1.AddHeader 60, RightJustify, "Column 3"
    ucReportList1.AddHeader 60, RightJustify, "Column 4"
    ucReportList1.AddHeader 60, RightJustify, "Column 5"
    
    For lItm = 1 To 100
        sRow = "Row " & lItm
        For lCol = 2 To ucReportList1.HeadersCount
            sRow = sRow & vbTab & Format$(10000 * Rnd, "#,0.00")
        Next lCol
        ucReportList1.AddItem sRow
    Next lItm
    ucReportList1.ListIndex = 0
    
' (*) You can set first column alignment to rigth by
'     adding as first column a zero-width column and
'     taking in account that you will need to start
'     from "second" tab.
End Sub

Private Sub Command1_Click()
  
  Dim sRow As String
  
    sRow = "Row changed"
    ucReportList1.List(ucReportList1.ListIndex) = sRow & vbTab & "-" & vbTab & "-" & vbTab & "-" & vbTab & "-"

' You can easily can access to sub-items through <Split>
End Sub
