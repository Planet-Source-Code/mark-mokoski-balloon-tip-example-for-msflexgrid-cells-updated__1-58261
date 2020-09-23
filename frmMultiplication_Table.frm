VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMultiplication_Table 
   BackColor       =   &H00800000&
   Caption         =   "Multiplication Table"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmMultiplication_Table.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7800
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "MS Flex Grid Control State"
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   5655
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         Picture         =   "frmMultiplication_Table.frx":1272
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSelRowCol 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtTipState 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Disabled"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtMouseRowCol 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtMouse_Y 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtMouse_X 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Row - Column"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Tip State"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Mouse Row - Column"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse ""Y"" Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse ""X"" Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000018&
      Caption         =   "Close Window"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      MouseIcon       =   "frmMultiplication_Table.frx":16B4
      MousePointer    =   99  'Custom
      Picture         =   "frmMultiplication_Table.frx":19BE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2740
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7865
      _ExtentX        =   13864
      _ExtentY        =   4842
      _Version        =   393216
      Rows            =   11
      Cols            =   11
      AllowBigSelection=   0   'False
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMultiplication_Table.frx":1E00
   End
   Begin VB.Label html2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7080
      MouseIcon       =   "frmMultiplication_Table.frx":211A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label html1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6240
      MouseIcon       =   "frmMultiplication_Table.frx":2424
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For the VB 6 Add-In version of the Balloon tip code generator, click"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   6600
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For the complete Balloon Tip code project, click"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label_Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press ""F11"" key or Right Click Mouse for Equation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balloon Tips on MSFlexgrid cell Example"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmMultiplication_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************
    '
    '   Test application for Sorrentino Massimiliano [sorrentino@ised.it]
    '   to test individual tool tips in a grid control
    '   using Mark Mokoski's ToolTip class
    '
    '   12-JAN-2005
    '   Mark Mokoski
    '
    '   2-FEB-2005
    '   Modified application for cell highlighting under mouse
    '   cursor. By request of another PSC member Arnold Donovan [newvbuser@yahoo.com]
    '   Most of the highlight code is in the MSFlexGrid1_MouseMove() event
    '
    '***************************************************

    Option Explicit
    
    'Shell out API for HTML files, Mail and Web Browser
    Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
    'API call to get Balloon Tips working
    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
    
    'Declare of a new Tool Tip Object
    'you need to have the clsToolTips in your project
    'For a more complete discription of the class
    'goto http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57489&lngWId=1
    'or click the links at the bottom of the application window
    
    Dim MSFlexGrid1_Tip                   As New clsTooltips
    
    Dim c                                 As Integer  'Column
    Dim r                                 As Integer  'Row
    Dim Operator1                         As String
    Dim Operator2                         As String
    Dim Result                            As String
    Dim Current_Row                       As Integer  'Current Row Mouse is over
    Dim Current_Col                       As Integer  'Current Column Mouse is over
    Dim Last_Row                          As Integer  'Last Row Mouse was over
    Dim Last_Col                          As Integer  'Last Column Mouse was over
    Dim LastSelected_Row                  As Integer
    Dim LastSelected_Col                  As Integer
    Dim TipText                           As String
    Dim New_Row                           As Integer
    Dim New_Col                           As Integer
    Dim MouseIn_grid                      As Boolean
    '"Mouse Over" Highlight Colors
    Dim SelectedCell_Forecolor            As Long
    Dim SelectedCell_Backcolor            As Long
    Dim Cell_Forecolor                    As Long
    Dim Cell_Backcolor                    As Long
    Dim Fixed_Forecolor                   As Long
    Dim Fixed_Backcolor                   As Long
    

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Call to API to get controls to work with Balloon Tips
    InitCommonControls
    
    'Turn control redraw OFF
    MSFlexGrid1.Redraw = False
    
    'Set at first fixed row
    MSFlexGrid1.Row = 0
    
    'Set column width and populate columns with numbers 1 thru 10

        For c = 0 To 10
            MSFlexGrid1.ColWidth(c) = 700
            MSFlexGrid1.Col = c
            MSFlexGrid1.CellAlignment = flexAlignCenterCenter
            
                
                If c = 0 Then   'Special case for Row(0), Col(0)
                    MSFlexGrid1.Text = "X"
                Else
                    MSFlexGrid1.Text = Str(c)
                End If

        Next c

    'Set first fixed column
    
    MSFlexGrid1.Col = 0
    
    'Polulate fixed rows with numbers 1 thru 10

        For r = 1 To 10
            MSFlexGrid1.Row = r
            MSFlexGrid1.CellAlignment = flexAlignCenterCenter
            MSFlexGrid1.Text = Str(r)
        Next r
        
        For r = 1 To 10
        
            'Polulate Grid with numbers ( c * r = celldata )

                For c = 1 To 10
                    MSFlexGrid1.Row = r
                    MSFlexGrid1.Col = c
                    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
                    MSFlexGrid1.Text = Str(r * c)
                Next c

        Next r

    'Set the default Hightlight colors
    Fixed_Forecolor = vbYellow
    Fixed_Backcolor = vbBlue
    SelectedCell_Forecolor = vbYellow
    SelectedCell_Backcolor = vbBlue
    Cell_Forecolor = vbWhite
    Cell_Backcolor = vbRed
    
    'Do some other form setup before main application starts
    
    'Set start posistion at row 1, column 1
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    
    MSFlexGrid1.Redraw = True
    
    'Set start values
    Current_Row = MSFlexGrid1.Row
    Last_Row = MSFlexGrid1.Row
    LastSelected_Row = MSFlexGrid1.Row
    Current_Col = MSFlexGrid1.Col
    Last_Col = MSFlexGrid1.Col
    LastSelected_Col = MSFlexGrid1.Col
    
    MSFlexGrid1_Click
    
End Sub

Private Sub html1_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57489&lngWId=1", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub html2_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute HWND, vbNullString, "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57698&lngWId=1", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub MSFlexGrid1_Click()

    'Set highlight color on fixed row/col of selected cell
    
    'Instead of keeping track of the last selected cell, you could
    'just brute force all the fixed rows / columns back to the
    'stock colors. But that might take too much time with a large grid.
    'So keeping track of the last selected cell is a minor inconvenience
    'in the hope of speeding up the grid redraw.
    
    'Turn off the redraw to stop flicker of grid control
    MSFlexGrid1.Redraw = False
    'Save the current selected cell
    New_Row = MSFlexGrid1.Row
    New_Col = MSFlexGrid1.Col
    
    'Reset color of last selected row fixed cell(s) and set color on new row

        For c = 0 To MSFlexGrid1.FixedCols - 1
            'Select and reset colors on old row
            MSFlexGrid1.Col = c
            MSFlexGrid1.Row = LastSelected_Row
            MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColorFixed
            MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColorFixed
            'select and set colors on new row
            MSFlexGrid1.Row = New_Row
            MSFlexGrid1.CellForeColor = Fixed_Forecolor
            MSFlexGrid1.CellBackColor = Fixed_Backcolor
        Next c
    
    'Reset color of last selected column fixed cell(s) and set color on new column

        For r = 0 To MSFlexGrid1.FixedRows - 1
            'Select and reset colors on old column
            MSFlexGrid1.Row = r
            MSFlexGrid1.Col = LastSelected_Col
            MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColorFixed
            MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColorFixed
            'select and set colors on new row
            MSFlexGrid1.Col = New_Col
            MSFlexGrid1.CellForeColor = Fixed_Forecolor
            MSFlexGrid1.CellBackColor = Fixed_Backcolor
        Next r
    
    'Restore currently selected cell
    MSFlexGrid1.Row = New_Row
    MSFlexGrid1.Col = New_Col
    
    'Redraw the control
    MSFlexGrid1.Redraw = True

    'Set cell current selection laso as last cell selected
    LastSelected_Row = MSFlexGrid1.Row
    LastSelected_Col = MSFlexGrid1.Col

    'On Mouse Click, set Row / Column text
    txtSelRowCol.Text = "Row " & Str(MSFlexGrid1.Row) & "  / Column " & Str(MSFlexGrid1.Col)
    DoEvents
  
    
End Sub

Private Sub MSFlexGrid1_GotFocus()

    'Turn off the redraw to stop flicker of grid control
    MSFlexGrid1.Redraw = False
    'Save the current selected cell
    New_Row = MSFlexGrid1.Row
    New_Col = MSFlexGrid1.Col
    
    'Set color of last selected row fixed cell(s) on control Got Focus

        For c = 0 To MSFlexGrid1.FixedCols - 1
            'Select and reset colors on old row
            MSFlexGrid1.Col = c
            MSFlexGrid1.Row = LastSelected_Row
            MSFlexGrid1.CellForeColor = Fixed_Forecolor
            MSFlexGrid1.CellBackColor = Fixed_Backcolor
        Next c
    
    'Set color of last selected column fixed cell(s) on control Got Focus

        For r = 0 To MSFlexGrid1.FixedRows - 1
            'Select and reset colors on old column
            MSFlexGrid1.Row = r
            MSFlexGrid1.Col = LastSelected_Col
            MSFlexGrid1.CellForeColor = Fixed_Forecolor
            MSFlexGrid1.CellBackColor = Fixed_Backcolor
        Next r
    
    'Restore currently selected cell
    MSFlexGrid1.Row = New_Row
    MSFlexGrid1.Col = New_Col
    
    'Redraw the control
    MSFlexGrid1.Redraw = True

    
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    'See if mouse is over the control

        If IsMouseOver(MSFlexGrid1) Then
        
            'If Key down on fixed row / columns, exit sub

                If MSFlexGrid1.MouseRow <= MSFlexGrid1.FixedRows - 1 Or MSFlexGrid1.MouseCol <= MSFlexGrid1.FixedCols - 1 Then Exit Sub

                Select Case KeyCode
                    Case vbKeyF11   'Keyboard "F11" key pressed
                        SetEquation 'Gosub to set cell selection and tool tip
                        
                End Select

        End If


    DoEvents

End Sub

Private Sub MSFlexGrid1_LostFocus()

    'If focus is lost to another control or window, remove the tool tip
    'Turn off the redraw to stop flicker of grid control
    MSFlexGrid1.Redraw = False
    'Save the current selected cell
    New_Row = MSFlexGrid1.Row
    New_Col = MSFlexGrid1.Col
    
    'Reset color of last selected row fixed cell(s) on control lost focus

        For c = 0 To MSFlexGrid1.FixedCols - 1
            'Select and reset colors on old row
            MSFlexGrid1.Col = c
            MSFlexGrid1.Row = LastSelected_Row
            MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColorFixed
            MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColorFixed
        Next c
    
    'Reset color of last selected column fixed cell(s) on control lost focus

        For r = 0 To MSFlexGrid1.FixedRows - 1
            'Select and reset colors on old column
            MSFlexGrid1.Row = r
            MSFlexGrid1.Col = LastSelected_Col
            MSFlexGrid1.CellForeColor = MSFlexGrid1.ForeColorFixed
            MSFlexGrid1.CellBackColor = MSFlexGrid1.BackColorFixed
        Next r
    
    'Restore currently selected cell
    MSFlexGrid1.Row = New_Row
    MSFlexGrid1.Col = New_Col
    
    'Redraw the control
    MSFlexGrid1.Redraw = True

    'Set some form display values
    MSFlexGrid1_Tip.Remove
    txtTipState.Text = "Disabled"
    Picture1.Visible = False
    DoEvents

End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    'See if mouse is over the control

        If IsMouseOver(MSFlexGrid1) Then

            'If mouse button down is on fixed row / colomn, exit sub

                If MSFlexGrid1.MouseRow <= MSFlexGrid1.FixedRows - 1 Or MSFlexGrid1.MouseCol <= MSFlexGrid1.FixedCols - 1 Then Exit Sub

                Select Case Button
                    Case vbKeyRButton   'Right Mouse button down
                        SetEquation     'Gosub to set cell selection and tool tip
                    Case vbKeyLButton   'Left Mouse button down
                        'Set cell selection without showing ToolTip
                        LastSelected_Row = MSFlexGrid1.Row
                        LastSelected_Col = MSFlexGrid1.Col
                        'Set the Highlight colors
                        MSFlexGrid1.CellForeColor = vbYellow
                        MSFlexGrid1.CellBackColor = vbBlue
               
                End Select

        End If

    DoEvents


End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'See where the mouse is, and if not in the "current" cell,
    'change the "current" cell data and remove the last tool tip

    Label_Info.Visible = True
    
    'Set some form display values
    txtMouse_X = Str(x)
    txtMouse_Y = Str(y)
   
   
        If MSFlexGrid1.MouseCol <> Current_Col Or MSFlexGrid1.MouseRow <> Current_Row Then
            'Save the cell that you came from
            Last_Row = Current_Row
            Last_Col = Current_Col
            
            'Now save the current Row,Column that the mouse is over
            Current_Row = MSFlexGrid1.MouseRow
            Current_Col = MSFlexGrid1.MouseCol
            
            'Turn off redraw property to stop "flicker" of selected cell
            MSFlexGrid1.Redraw = False

            'Save Last selected Row/Column
            LastSelected_Row = MSFlexGrid1.Row
            LastSelected_Col = MSFlexGrid1.Col
            'Reselect the row,column
            MSFlexGrid1.Row = Last_Row
            MSFlexGrid1.Col = Last_Col
            DoEvents
            
            'If row or column is the fixed row/column, do not change color

                If Last_Row <> MSFlexGrid1.FixedRows - 1 And Last_Col <> MSFlexGrid1.FixedCols - 1 Then
                    'Reset cell colors
                    MSFlexGrid1.CellForeColor = vbBlack
                    MSFlexGrid1.CellBackColor = vbWhite
                End If
            
            'Now reselect the cell the mouse is over
            MSFlexGrid1.Row = Current_Row
            MSFlexGrid1.Col = Current_Col
            DoEvents
            
            'If row or column is the fixed row/column, do not change color

                If Current_Row <> MSFlexGrid1.FixedRows - 1 And Current_Col <> MSFlexGrid1.FixedCols - 1 Then
                    'Change the current cells "Highlighted" colors
                    
                    'If mouse over current selected cell, usr these colors

                        If MSFlexGrid1.Row = LastSelected_Row And MSFlexGrid1.Col = LastSelected_Col Then
                            MSFlexGrid1.CellForeColor = vbYellow
                            MSFlexGrid1.CellBackColor = vbBlue
                        Else
                            'If mouse over any other call, use tese colors
                            MSFlexGrid1.CellForeColor = vbWhite
                            MSFlexGrid1.CellBackColor = vbRed
                        End If

                End If

            'Now restore the REAL last Mouse or cursor selected column
            MSFlexGrid1.Row = LastSelected_Row
            MSFlexGrid1.Col = LastSelected_Col
            DoEvents
            
            
            'Kill any tool tips
            MSFlexGrid1_Tip.Remove

            'Turn redraw property "on" to refelect color changes
            MSFlexGrid1.Redraw = True
            DoEvents
            
            'Set some form display values
            txtMouseRowCol.Text = "Row " & Str(MSFlexGrid1.MouseRow) & "  / Column " & Str(MSFlexGrid1.MouseCol)
            txtTipState.Text = "Disabled"
            Picture1.Visible = False
            
        End If
        
End Sub

Private Sub MSFlexGrid1_RowColChange()
    
    'Disable (remove) Tool Tip if selected cell has changed

        If MSFlexGrid1.Redraw = True Then
            'Remove any active tool tip
            MSFlexGrid1_Tip.Remove
    
            'Set some form display values
            txtTipState.Text = "Disabled"
            Picture1.Visible = False
            txtSelRowCol.Text = "Row " & Str(MSFlexGrid1.Row) & "  / Column " & Str(MSFlexGrid1.Col)
    
            MSFlexGrid1_Click
        End If

End Sub

Private Sub Timer1_Timer()

    'Timer for "Mouse_Over" event, see modMouseOver
    'I know it's a bit of a cheat using a timer,
    'but just for checking for the "Mouse Over" on just one control
    'I didn't see the need for subclassing or hooks just for this
    'type project. See modMouseOver for code.
    
    On Error Resume Next

        If Not IsMouseOver(MSFlexGrid1) And MouseIn_grid = True Then
            txtMouse_X.Text = ""
            txtMouse_Y.Text = ""
            txtMouseRowCol.Text = ""
            MSFlexGrid1_Tip.Remove
            txtTipState.Text = "Disabled"
            Picture1.Visible = False
            Label_Info.Visible = False
            
            MSFlexGrid1.Redraw = False
            
            MSFlexGrid1.Row = Current_Row
            MSFlexGrid1.Col = Current_Col
            DoEvents
            
            'If row or column is the fixed row/column, do not change color

                If Current_Row <> MSFlexGrid1.FixedRows - 1 And Current_Col <> MSFlexGrid1.FixedCols - 1 Then
                    'Reset cell colors
                    MSFlexGrid1.CellForeColor = vbBlack
                    MSFlexGrid1.CellBackColor = vbWhite
                End If
                
            MSFlexGrid1.Row = LastSelected_Row
            MSFlexGrid1.Col = LastSelected_Col
            DoEvents
            
            Current_Row = 0
            Current_Col = 0
            
            MSFlexGrid1.Redraw = True
        
        End If

        'If mouse enters the control without focus, set the focus to the control
        If IsMouseOver(MSFlexGrid1) And MouseIn_grid = False Then
            MSFlexGrid1.SetFocus
        End If

    MouseIn_grid = IsMouseOver(MSFlexGrid1)

End Sub

Private Sub SetEquation()

    'Get the data to for an equation. Format x * y = z
    'and setup Tool Tip to show for that cell
    MSFlexGrid1.Redraw = False
    
    MSFlexGrid1.Col = MSFlexGrid1.MouseCol
    MSFlexGrid1.Row = MSFlexGrid1.MouseRow
    MSFlexGrid1_Click
    Result = MSFlexGrid1.Text
    MSFlexGrid1.Col = 0
    Operator1 = MSFlexGrid1.Text
    MSFlexGrid1.Col = MSFlexGrid1.MouseCol
    MSFlexGrid1.Row = 0
    Operator2 = MSFlexGrid1.Text
    MSFlexGrid1.Row = MSFlexGrid1.MouseRow
    
    'Build the Balloon Tool Tip object
    TipText = "Equation is " & Operator1 & "  X " & Operator2 & " = " & Result
    MSFlexGrid1_Tip.CreateBalloon MSFlexGrid1, _
    TipText, _
    "Multiplication Table", 1
    
    'Set some form infomation
    txtTipState.Text = "Enabled"
    Picture1.Visible = True
    
    'Set the Highlight colors
    MSFlexGrid1.CellForeColor = vbYellow
    MSFlexGrid1.CellBackColor = vbBlue
    
    'CC_Comment Out (2/6/2005):
    '    'Set current cell as last selected
    '    LastSelected_Row = MSFlexGrid1.Row
    '    LastSelected_Col = MSFlexGrid1.Col

    'End CC_Comment Out
    MSFlexGrid1.Row = LastSelected_Row
    MSFlexGrid1.Col = LastSelected_Col

    MSFlexGrid1.Redraw = True

End Sub
