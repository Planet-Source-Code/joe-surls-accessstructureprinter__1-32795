VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Database Structure Print Utility"
   ClientHeight    =   3465
   ClientLeft      =   2565
   ClientTop       =   2565
   ClientWidth     =   5190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   3465
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   90
      TabIndex        =   6
      Top             =   1560
      Width           =   5010
      Begin VB.CheckBox chkSeparated 
         Caption         =   "Separate Page Per Table"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkSystemTables 
         Caption         =   "Print System Tables"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "HTML"
         Height          =   195
         Left            =   3045
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optPrinter 
         Caption         =   "Printer"
         Height          =   195
         Left            =   3045
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   1688
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   825
      Top             =   2775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Access Database"
      Filter          =   "Access Databases *.mdb |*.mdb"
      InitDir         =   "C:\"
   End
   Begin VB.TextBox txtDBPath 
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   128
      TabIndex        =   1
      Top             =   960
      Width           =   3870
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   4028
      TabIndex        =   0
      Top             =   960
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4950
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 - Click Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   878
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Set Your Print Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   878
      TabIndex        =   3
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - Select Your Access Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   878
      TabIndex        =   2
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code Author : Joseph B. Surls
'Author E-Mail: joseph.surls@verizon.net
'Submission Date : March 17, 2002       Happy St. Pat's!!

'I wrote this code for work. When you don't have Access
'on a user's machine where your VB program is, you can
'still check out its structure to pinpoint a problem.

'I also have a "Database Helper" program that i wrote
'that allows you to view teh tables in the DB, run custom
'Select queries and Executes against an Access DB.
'Maybe I'll submit that one next.

'I hope this code helps somebody. Feel free to use this code
'and/or change it in any way. Drop me an E and let me know
'if I can help in any way.

'I'm thinking about adding the ability to print Queries adn
'generally tweaking it up a bit. Maybe if I get some good feedback
'it'll motivate me a little.

'Thx, Joe

Option Explicit
'Database Object
Dim dbAccess As DAO.Database
'Recordset Object
Dim rsAccess As DAO.Recordset
Dim i As Integer
Dim j As Long
'TableDef Object
Dim oTable As DAO.TableDef
'Field Object
Dim oField As DAO.Field

Private Sub cmdBrowse_Click()
On Error GoTo CancelBrowse
    'Open Common Dialog for User to input Database Path
    With dlgCommon
        .CancelError = True
        .InitDir = App.Path
        .DialogTitle = "Open Database..."
        .Filter = "Access Databases *.mdb|*.mdb"
        .FileName = ""
        .ShowOpen
        txtDBPath = .FileName
    End With
    Exit Sub
CancelBrowse:
    If Err.Number = 32755 Then 'User Pressed Cancel button
        Exit Sub
    Else
        MsgBox Err.Number & Chr(10) & _
            Err.Description
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo NoDB
    'If no printers on user's system, get out
    If Printers.Count < 1 Then Exit Sub
    
    'If no DB specified, get out
    If txtDBPath = "" Then Exit Sub
    
    'this is for Password-protected Access databases
    If frmPassword.pstrPassword = "" Then
        'No password (if password-protected, will error out
        'and show "Enter Password" form
        Set dbAccess = OpenDatabase(Trim(txtDBPath), True, True)
    Else
        'Password has been specified
        Set dbAccess = OpenDatabase(Trim(txtDBPath), True, True, ";pwd=" & frmPassword.pstrPassword)
        frmPassword.pstrPassword = ""
    End If
    
    If optHTML.Value = True Then 'Print Structure in an HTML file
        PrintHTML
        Set dbAccess = Nothing
        Exit Sub
    Else                         'Print Structure to a printer
        Screen.MousePointer = vbHourglass
        Printer.Print Trim(txtDBPath)
        Printer.Print ""
        Printer.Print ""
        For Each oTable In dbAccess.TableDefs 'Loop through each table in the database
        
            'this next line determines whether to print the Access System tables or not
            If chkSystemTables.Value = vbChecked Or Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
                
                'Printer Setup Header Stuff
                Printer.FontSize = 14
                Printer.FontBold = True
                Printer.Print "TABLE NAME = " & oTable.Name
                Printer.FontSize = 8
                Printer.FontBold = False
                Printer.Print "======================================="
                Printer.Print "Date Created =" & oTable.DateCreated
                Printer.Print "Date Last Modified = " & oTable.LastUpdated
                Printer.Print "Records = " & oTable.RecordCount
                Printer.Print "---------------------------------------------------"
                Printer.Print ""
                Printer.Print ""
                
                'Dont print System table breakdown
                If Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
                    'open recordset on current table
                    Set rsAccess = dbAccess.OpenRecordset(oTable.Name, dbOpenTable)
                    
                    'All this X and Y stuff sets up the Columns and headings
                    Printer.CurrentX = 500
                    Printer.FontBold = True
                    Printer.Print "Fields Listing"
                    Printer.FontBold = False
                    Printer.CurrentX = 1000
                    j = Printer.CurrentY
                    Printer.Print "Field Name"
                    Printer.CurrentX = 3000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "Type"
                    Printer.CurrentX = 5000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "Size"
                    Printer.CurrentX = 7000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "Required"
                    Printer.CurrentX = 9000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "Allow Null"
                    Printer.CurrentX = 1000
                    j = Printer.CurrentY
                    Printer.Print "-------------------"
                    Printer.CurrentX = 3000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "--------"
                    Printer.CurrentX = 5000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "--------"
                    Printer.CurrentX = 7000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "--------------"
                    Printer.CurrentX = 9000
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.Print "---------------"
                    i = 0
                    
                    'Loop thru each field in current table
                    'Line up columns and print field info
                    For Each oField In rsAccess.Fields
                        Printer.CurrentX = 1000
                        j = Printer.CurrentY
                        Printer.Print oField.Name
                        Printer.CurrentX = 3000
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        
                        'convert datatype into English
                        Printer.Print GetFieldType(oField.Type)
                        
                        Printer.CurrentX = 5000
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        Printer.Print oField.Size
                        Printer.CurrentX = 7000
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        Printer.Print oField.Required
                        Printer.CurrentX = 9000
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        Printer.Print oField.AllowZeroLength
                        i = i + 1
                    Next
                End If
                
                'Get any indexes for current table
                If oTable.Indexes.Count > 0 Then
                    Printer.Print ""
                    Printer.CurrentX = 500
                    Printer.FontBold = True
                    Printer.Print "Index Listing"
                    Printer.FontBold = False
                    j = Printer.CurrentY
                    Printer.CurrentX = 1000
                    Printer.Print "Index Name"
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.CurrentX = 3000
                    Printer.Print "Fields"
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.CurrentX = 6000
                    Printer.Print "Unique"
                    j = Printer.CurrentY
                    Printer.CurrentX = 1000
                    Printer.Print "----------------"
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.CurrentX = 3000
                    Printer.Print "----------"
                    If Printer.CurrentY < j Then
                        j = Printer.CurrentY
                    End If
                    Printer.CurrentY = j
                    Printer.CurrentX = 6000
                    Printer.Print "----------"
                    
                    'loop thru table Indexes (if any)
                    For i = 0 To oTable.Indexes.Count - 1
                        j = Printer.CurrentY
                        Printer.CurrentX = 1000
                        Printer.Print oTable.Indexes(i).Name
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        Printer.CurrentX = 3000
                        Printer.Print oTable.Indexes(i).Fields
                        If Printer.CurrentY < j Then
                            j = Printer.CurrentY
                        End If
                        Printer.CurrentY = j
                        Printer.CurrentX = 6000
                        Printer.Print oTable.Indexes(i).Unique
                    Next i
                End If
                
                'Clear recordset for next table
                Set rsAccess = Nothing
                
                'Print each table on separate page or not
                If chkSeparated.Value = vbChecked Then
                    Printer.EndDoc
                Else
                    Printer.Print ""
                    Printer.Print ""
                End If
            End If
        Next
        If Not chkSeparated.Value = vbChecked Then
            Printer.EndDoc
        End If
        
        'Clear database variable
        Set dbAccess = Nothing
        MsgBox "Your Access Structure has been printed to " & Printer.DeviceName, vbInformation, "Complete"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
NoDB:
    If Err.Number = 3031 Then 'Database needs a password
        frmPassword.Show vbModal
        If frmPassword.pblnCancel = True Then Exit Sub
        cmdPrint_Click
        Err.Clear
        Exit Sub
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Function GetFieldType(TypeCode As Integer)
'This routine accepts the Fieldtype variable and returns the
'English version for printing
    Select Case TypeCode
        Case dbBinary
            GetFieldType = "Binary"
        Case dbBoolean
            GetFieldType = "Boolean"
        Case dbByte
            GetFieldType = "Byte"
        Case dbChar
            GetFieldType = "Character"
        Case dbCurrency
            GetFieldType = "Currency"
        Case dbDate
            GetFieldType = "Date/Time"
        Case dbDecimal
            GetFieldType = "Decimal"
        Case dbDouble
            GetFieldType = "Double"
        Case dbFloat
            GetFieldType = "Float"
        Case dbGUID
            GetFieldType = "GUID"
        Case dbInteger
            GetFieldType = "Integer"
        Case dbLong
            GetFieldType = "Long"
        Case dbLongBinary
            GetFieldType = "OLE Object"
        Case dbMemo
            GetFieldType = "Memo"
        Case dbNumeric
            GetFieldType = "Numeric"
        Case dbSingle
            GetFieldType = "Single"
        Case dbText
            GetFieldType = "Text"
        Case dbTime
            GetFieldType = "Time"
        Case dbTimeStamp
            GetFieldType = "TimeStamp"
        Case dbVarBinary
            GetFieldType = "VarBinary"
        Case Else
            GetFieldType = "Undetermined"
    End Select
End Function

Private Sub PrintHTML()
'this routine prints the Access Structure to an HTML file
Dim SaveFile As String
On Error GoTo CancelHTML
    
    'More Common Dialog
    With dlgCommon
        .CancelError = True
        .DialogTitle = "Save HTML Page As..."
        .Filter = "Web Page *.htm|*.htm;*.html"
        .InitDir = "C:\"
        .FileName = "Structure.htm"
        .ShowSave
        SaveFile = .FileName
    End With
    DoEvents
    Open SaveFile For Output As #2
    
    'Set database Object
    Set dbAccess = OpenDatabase(Trim(txtDBPath), True, True)
    
    'HTML Template stuff
    Print #2, "<html>"
    Print #2, "<head>"
    Print #2, "<meta name='Access Structure Print' content=Joseph Surls'>"
    Print #2, "<title>" & "Access Structure for " & Trim(txtDBPath) & "</title>"
    Print #2, "</head>"
    Print #2, "<body bgcolor='#0099FF'>"
    Print #2, "<p><font size='1'>"
    Print #2, Trim(txtDBPath)
    Print #2, "</a></font></p>"
    
    'Loop thru each table in Database
    For Each oTable In dbAccess.TableDefs
        Print #2, "<p><b><u><font size='4' color='#000000'>"
        Print #2, "Table " & oTable.Name & "</font><br>"
        Print #2, "</u></b><font size='2'>"
        Print #2, "Date Created - " & oTable.DateCreated & "<br>"
        Print #2, "Date Last Modified - " & oTable.LastUpdated & "<br>"
        Print #2, "Records - " & oTable.RecordCount & "<br>"
        Print #2, "-----------------------------------------------------------"
        Print #2, "</font></p>"
        
        'No System Tables
        If Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
            
            'open recordset for each table
            Set rsAccess = dbAccess.OpenRecordset(oTable.Name, dbOpenTable)
            Print #2, "<p>&nbsp;&nbsp; <font size='2'> </font><b><font size='3'>Fields Listing</font></b></p>"
            Print #2, "<table border='0' width='100%'>"
            Print #2, "<tr><td width='10%' align='center'></td>"
            Print #2, "<td width='30%' align='center'>"
            Print #2, "<p align='center'><u>Field Name</u></td>"
            Print #2, "<td width='20%' align='center'><u>Type</u></td>"
            Print #2, "<td width='10%' align='center'><u>Size</u></td>"
            Print #2, "<td width='10%' align='center'><u>Required</u></td>"
            Print #2, "<td width='44%' align='center'><u>Allow Null</u></td>"
            Print #2, "<td width='16%' align='center'></td></tr>"
            
            'Loop thru each field in current table
            For Each oField In rsAccess.Fields
                Print #2, "<tr><td width='10%' align='center'></td>"
                Print #2, "<td width='30%' align='center'>"
                Print #2, oField.Name & "</td>"
                Print #2, "<td width='20%' align='center'>"
                
                'convert data type to English
                Print #2, GetFieldType(oField.Type) & "</td>"
                Print #2, "<td width='10%' align='center'>"
                Print #2, oField.Size & "</td>"
                Print #2, "<td width='10%' align='center'>"
                Print #2, oField.Required & "</td>"
                Print #2, "<td width='44%' align='center'>"
                Print #2, oField.AllowZeroLength & "</td>"
                Print #2, "<td width='16%' align='center'></td>"
                Print #2, "</tr>"
            Next
            Print #2, "</table>"
            
            'Table Indexes
            If oTable.Indexes.Count > 0 Then
                Print #2, "<p>&nbsp;&nbsp;&nbsp; <b>Index Listing</b></p>"
                Print #2, "<table border='0' width='100%'>"
                Print #2, "<tr>"
                Print #2, "<td width='7%' align='center'></td>"
                Print #2, "<td width='23%' align='center'><u>Index Name</u></td>"
                Print #2, "<td width='44%' align='center'><u>Fields</u></td>"
                Print #2, "<td width='19%' align='center'><u>Unique</u></td>"
                Print #2, "<td width='7%' align='center'></td>"
                Print #2, "</tr>"
                For i = 0 To oTable.Indexes.Count - 1
                    Print #2, "<tr>"
                    Print #2, "<td width='7%' align='center'></td>"
                    Print #2, "<td width='23%' align='center'>"
                    Print #2, oTable.Indexes(i).Name & "</td>"
                    Print #2, "<td width='44%' align='center'>"
                    Print #2, oTable.Indexes(i).Fields & "</td>"
                    Print #2, "<td width='19%' align='center'>"
                    Print #2, oTable.Indexes(i).Unique & "</td>"
                    Print #2, "<td width='7%' align='center'></td>"
                    Print #2, "</tr>"
                Next i
            End If
            Print #2, "</table>"
            Print #2, "<p>====================================================================================</p>"
        End If
    Next
    Print #2, "<p align='center'>End of Listing<br>"
    Print #2, "This Page Created by Access Structure Print Software - " & _
        Date & "</p>"
    Print #2, "</body>"
    Print #2, "</html>"
    Close #2
    MsgBox "Your HTML Listing has been saved as " & dlgCommon.FileName, vbInformation, "Complete"
    Exit Sub
CancelHTML:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox Err.Number & Chr(10) & _
            Err.Description
    End If
End Sub

Private Sub optHTML_Click()
    'Disable Irrelevant Check Buttons
    chkSeparated.Enabled = False
    chkSystemTables.Enabled = False
End Sub

Private Sub optPrinter_Click()
    'Enable Relevant Check Buttons
    chkSeparated.Enabled = True
    chkSystemTables.Enabled = True
End Sub
