VERSION 5.00
Begin VB.Form frmMainChart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Reports/Chart Director Report"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmMainChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comboReport 
      Height          =   315
      ItemData        =   "frmMainChart.frx":000C
      Left            =   150
      List            =   "frmMainChart.frx":0013
      TabIndex        =   9
      Top             =   540
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output Details"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1575
      Width           =   8220
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optUndefined 
         Caption         =   "Undefined"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optPrinter 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optScreen 
         Caption         =   "Screen"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Report"
      Height          =   300
      Left            =   2805
      TabIndex        =   2
      Top             =   555
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   300
      Left            =   4305
      TabIndex        =   1
      Top             =   555
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmMainChart.frx":0022
      Height          =   495
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   930
      Width           =   6735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Visual Reports/Chart Director Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   8265
   End
End
Attribute VB_Name = "frmMainChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------
Rem Application :   Visual Reports/Chart Director Report
Rem Description :   Produce a report using Visual Reports
Rem                 and Chart Director
Rem
Rem Dependents  :   Developer will need Visual Reports
Rem                 and Chart Director to run this example
Rem ---------------------------------------------------
Option Explicit

'The main VisualReport object
Private WithEvents m_objReport As Report
Attribute m_objReport.VB_VarHelpID = -1
'Used by report writer for the data
Dim m_db As DAO.Database
Dim m_ws As DAO.Workspace

'Used to track previous report selection
Dim m_nPrevIndex As Integer

Private Sub cmdExit_Click()
    'clean up before closing
    
    If Not m_objReport Is Nothing Then
        Set m_objReport = Nothing
    End If
 
    End
End Sub

Private Sub cmdRun_Click()
    Me.MousePointer = vbHourglass
    Select Case comboReport.ItemData(comboReport.ListIndex)
      Case 1
           If InitialiseReport("cdpie.vrd") Then
              PrintCDPieReport
           End If
    End Select
    Me.MousePointer = vbDefault
    
End Sub


Private Sub comboReport_Click()
    lblDescription(m_nPrevIndex).Visible = False
    lblDescription(comboReport.ListIndex).Visible = True
    lblDescription(comboReport.ListIndex).Left = lblDescription(0).Left
    lblDescription(comboReport.ListIndex).Top = lblDescription(0).Top
    m_nPrevIndex = comboReport.ListIndex
End Sub

Private Sub Form_Load()
    Dim objPrinter As Printer
    
    cboPrinter.AddItem "<Undefined>"
    For Each objPrinter In Printers
        cboPrinter.AddItem objPrinter.DeviceName
    Next objPrinter
    cboPrinter.ListIndex = 0
    comboReport.ListIndex = 0
    m_nPrevIndex = 0
End Sub



Private Sub m_objReport_NewPage()
    'The Visual Report object fires this
    'event every time it starts a new page,
    'except for the first page which gives
    'you a chance to print a page header.
    Dim nCount As Integer
    For nCount = 1 To m_objReport.SectionCount
        If LCase(m_objReport.Section(nCount).Name) = "pageheader" Then
            m_objReport.Section("PageHeader").Field("date").Value = Format(Now, "mm/dd/yyyy")
            m_objReport.PrintSection "PageHeader"
            Exit Sub
        End If
    Next nCount
End Sub
Private Sub m_objReport_PrintFooter()
    'The Visual Report object fires this
    'event before an end a page occur,
    'you get a chance to print a page footer.
    Dim nCount As Integer
    For nCount = 1 To m_objReport.SectionCount
        If LCase(m_objReport.Section(nCount).Name) = "pagefooter" Then
            m_objReport.Section("PageFooter").Field("PageNumber").Value = m_objReport.Section("PageFooter").Field("PageNumber").Value + 1
            m_objReport.PrintSection "PageFooter"
            Exit Sub
        End If
    Next nCount
End Sub

Private Sub Report_PrintToPrinterButton()
    'If the user presses the print to printer
    'button after first printing to screen, this
    'event is fired.  To print to printer, the output is changed to printer, and the report
    'is printed all over again.
    m_objReport.OutputTo = cOutputToPrinter
    cmdRun_Click
End Sub

Private Function InitialiseReport(sReportName As String) As Boolean
    'This functions is used to open the report
    'file and initialise the report object.
    On Error GoTo Err_handler
    InitialiseReport = False
    
    If Not m_objReport Is Nothing Then
        Set m_objReport = Nothing
    End If
    Set m_objReport = New Report
    Screen.MousePointer = vbHourglass
    
    If m_objReport.LoadReport(App.Path & "\" & sReportName) Then
        InitialiseReport = True
    End If
    
    GetOutput
    Screen.MousePointer = vbDefault
    Exit Function
    
Err_handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbOKOnly, Me.Name & "-InitialiseReport"
    m_objReport.EndReport
    Exit Function
End Function

Private Sub OpenDatabase()
    'Opens the access database using DAO
    On Error GoTo Err_handler
    If Not m_db Is Nothing Then m_db.Close
    Set m_ws = DBEngine.Workspaces(0)
    Set m_db = m_ws.OpenDatabase(App.Path & "\Nwlite.mdb")
    Exit Sub
Err_handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbOKOnly, Me.Name & "-InitialiseReport"
    m_objReport.EndReport
    
    
End Sub

Private Sub CloseDatabase()
    'Closes the access database
    If Not m_db Is Nothing Then
        m_db.Close
        Set m_db = Nothing
    End If
End Sub


Private Sub GetOutput()
    m_objReport.PrinterName = ""
    If optUndefined.Value Then m_objReport.OutputTo = cOutputToUnknown
    If optScreen.Value Then m_objReport.OutputTo = cOutputToScreen
    If optPrinter.Value Then m_objReport.OutputTo = cOutputToPrinter
    If cboPrinter.Text <> "" And cboPrinter.Text <> "<Undefined>" Then m_objReport.PrinterName = cboPrinter.Text
End Sub


Private Sub PrintCDPieReport()
    'Prints the Pie Chart report

    Dim rs As DAO.Recordset
    Dim p As New Pie
    Dim pi As PieItem
    Dim arColor(9) As Long
    Dim nColorIndex As Integer
    Dim nSecY As Double
    Dim nSecX As Double
    Dim data() As Variant
    Dim labels() As Variant
    Dim nDataIndex As Integer
    
    OpenDatabase
    
    m_objReport.WindowState = cMaximized

    Dim cd As Object
    Set cd = CreateObject("ChartDirector.API")
    
    'First, create a PieChart of size 360 pixels x 300 pixels
    Dim c As Object
    Set c = cd.PieChart(360, 300)
    
    'Set the center of the pie at (180, 140) and the radius to 100 pixels
    Call c.setPieSize(110, 140, 70)
        
    'Print Header
    m_objReport.Section("PageHeader").Field("date").Value = Format(Now, "mm/dd/yyyy")
    m_objReport.Section("PageHeader").Field("northwind_logo").Value = App.Path & "\northwind2.bmp"
    m_objReport.PrintSection "PageHeader"     'and page header sections.
    
    Set rs = m_db.OpenRecordset("SELECT Employees.LastName & "", "" & Employees.FirstName AS EmployeeName, Count(*) AS OrderCount FROM Orders, Employees WHERE (((DatePart(""yyyy"", [Orders].[OrderDate])) = 1995) And ((Orders.EmployeeID) = [Employees].[EmployeeID]))GROUP BY Employees.LastName & "", "" & Employees.FirstName")
    
    nSecY = m_objReport.CorrY
    m_objReport.CorrY = m_objReport.CorrY + 500
    m_objReport.PrintSection ("EmployeeHeader")
    nDataIndex = 0
    ReDim data(10)
    ReDim labels(10)
    
    Do Until rs.EOF
       data(nDataIndex) = rs("OrderCount")
       labels(nDataIndex) = rs("EmployeeName")
       With m_objReport.Section("EmployeeDetail")
         .Clear
         .Field("EmployeeName").Value = rs("EmployeeName")
         .Field("OrderCount").Value = rs("OrderCount")
       End With
       m_objReport.PrintSection ("EmployeeDetail")
       nDataIndex = nDataIndex + 1
       rs.MoveNext
    Loop
    m_objReport.CorrY = nSecY
    
    'Set the pie data and the pie labels
    ReDim Preserve data(nDataIndex - 1)
    ReDim Preserve labels(nDataIndex - 1)
    
    'set chart data
    Call c.SetData(data, labels)
    
    Call c.setLabelFormat("&percent&%")

    'add legends
    Call c.addLegend(240, 60, 1, "arial.ttf", 8).setBackground(RGB(255, 255, 255), RGB(0, 0, 0))
    'output the chart
    Call c.makeChart(App.Path & "\image0.gif")
    m_objReport.Section("PieSection").Field("pie_1").Value = App.Path & "\image0.gif"
    m_objReport.PrintSection ("PieSection")
    
end_rpt:
    m_objReport.EndReport                             'Finish off the report.
    CloseDatabase
    Set cd = Nothing
End Sub
