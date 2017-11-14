VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPrint7.ocx"
Begin VB.Form Form1 
   Caption         =   "VSPrinter7: Improved Tables"
   ClientHeight    =   7920
   ClientLeft      =   2160
   ClientTop       =   1770
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8160
   Begin VB.CheckBox chkCustomBorders 
      Caption         =   "&Custom Borders"
      Height          =   300
      Left            =   90
      TabIndex        =   10
      Top             =   675
      Width           =   1800
   End
   Begin VB.CheckBox chkOwnerDraw 
      Caption         =   "&OwnerDraw"
      Height          =   300
      Left            =   3465
      TabIndex        =   8
      Top             =   945
      Width           =   1800
   End
   Begin VB.CheckBox chkVerticalText 
      Caption         =   "&Vertical Text"
      Height          =   300
      Left            =   3465
      TabIndex        =   6
      Top             =   675
      Width           =   1800
   End
   Begin VB.CheckBox chkRowSpan 
      Caption         =   "&RowSpan"
      Height          =   300
      Left            =   5265
      TabIndex        =   5
      Top             =   675
      Width           =   1305
   End
   Begin VB.CheckBox chkColSpan 
      Caption         =   "&ColSpan"
      Height          =   300
      Left            =   5265
      TabIndex        =   4
      Top             =   405
      Width           =   1305
   End
   Begin VB.CommandButton btnRender 
      Caption         =   "Render Table"
      Default         =   -1  'True
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   990
      Width           =   3120
   End
   Begin VB.ComboBox cmbBorders 
      Height          =   315
      ItemData        =   "vptable.frx":0000
      Left            =   90
      List            =   "vptable.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   3120
   End
   Begin VB.CheckBox chkKeepTogether 
      Caption         =   "&Keep rows together"
      Height          =   300
      Left            =   3465
      TabIndex        =   1
      Top             =   405
      Width           =   1800
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Align           =   2  'Align Bottom
      Height          =   6480
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   8160
      _cx             =   14393
      _cy             =   11430
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   -1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   35.8901515151515
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3330
      X2              =   3330
      Y1              =   90
      Y2              =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Table Border"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   3120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Special Effects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3465
      TabIndex        =   7
      Top             =   90
      Width           =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    cmbBorders.ListIndex = 7 ' tbAll
End Sub

Private Sub Form_Resize()
    'On Error Resume Next
    vp.Height = ScaleHeight - (btnRender.Top + btnRender.Height + 100)
End Sub

Private Sub btnRender_Click()
    
    ' create a long string to add to some table cells
    Dim sLong$
    sLong = "This is the very long string that will cause some rows to break across pages. "
    sLong = sLong & vbCrLf & sLong & vbCrLf & sLong
    
    ' create document
    With vp
        .StartDoc
        
        ' show intro
        .Paragraph = "VSPrinter7 has more powerful tables than previous versions. " & vbCrLf & _
                     "New features include breaking rows across pages, vertical text " & _
                     "in cells, row spanning, customizable borders, and better events " & _
                     "to allow custom painting." & vbCrLf & vbCrLf
                     
        ' set page and table borders
        .PageBorder = pbAll
        .TableBorder = Val(cmbBorders)
        .TablePenLR = 40
        .TablePenTB = 40
        
        ' build table with 10 rows
        .StartTable
        .AddTable "2300|2300|2300", "Column 1|Column 2|Column 3", "", RGB(200, 200, 250)
        .TableCell(tcRows) = 10
        
        ' center align all cells
        .TableCell(tcAlign) = taCenterMiddle
                
        ' add text to all cells
        Dim row%, col%
        For row = 1 To 10
            For col = 1 To 3
                If (row + col) Mod 7 <> 0 Then
                    .TableCell(tcText, row, col) = " Row " & row & " Col " & col & " "
                Else
                    ' make a few cells have longer text, bold with a background
                    .TableCell(tcText, row, col) = sLong
                    .TableCell(tcBackColor, row, col) = RGB(100, 250, 100)
                    .TableCell(tcFontBold, row, col) = True
                End If
            Next
        Next
        
        ' keep rows together
        .TableCell(tcRowKeepTogether) = chkKeepTogether.Value
        
        ' apply vertical text to first row
        If chkVerticalText.Value Then
            .TableCell(tcVertical, 1, 1, 1, 3) = True
        End If
                        
        ' apply colspan
        If chkColSpan.Value Then
            .TableCell(tcColSpan, 1, 1) = 2
            .TableCell(tcBackColor, 1, 1) = RGB(250, 100, 100)
        End If
                
        ' apply rowspan
        If chkRowSpan.Value Then
            .TableCell(tcRowSpan, 3, 2) = 2
            .TableCell(tcBackColor, 3, 2) = RGB(250, 100, 100)
        End If
                
        ' apply custom borders
        If chkCustomBorders.Value Then
            .TableCell(tcRowBorderAbove, 2) = 80
            .TableCell(tcRowBorderBelow, 2) = 80
            .TableCell(tcRowBorderColor, 2) = RGB(0, 0, 100)
            .TableCell(tcColBorderLeft, , 2) = 40
            .TableCell(tcColBorderRight, , 2) = 40
            .TableCell(tcColBorderColor, , 2) = RGB(0, 100, 0)
        End If
        
        .EndTable
        .EndDoc
    End With
End Sub

Private Sub vp_AfterTableCell(ByVal row As Long, ByVal col As Long, ByVal Left As Double, ByVal Top As Double, ByVal Right As Double, ByVal Bottom As Double, Text As String, KeepFiring As Boolean)

    ' draw a cross over the cell
    If chkOwnerDraw.Value <> 0 Then
        KeepFiring = True
        If Len(Text) > 100 Then
            vp.DrawLine Left, Top, Right, Bottom
            vp.DrawLine Left, Bottom, Right, Top
        End If
    End If
End Sub
