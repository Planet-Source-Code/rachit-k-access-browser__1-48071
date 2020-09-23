VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFCEC&
   Caption         =   "Access Browser"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton CmdGo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFCEC&
      Caption         =   "BROWSE TABLE"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CmdLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< &Back"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   2655
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4683
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorBkg    =   -2147483632
         GridColorFixed  =   8421376
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7200
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   7200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFCEC&
      Caption         =   "Authentication Required"
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFCEC&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   3600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   3600
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   2640
         X2              =   4800
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFCEC&
      Caption         =   "Field Names"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "AccessBrowser.frx":0000
         Left            =   120
         List            =   "AccessBrowser.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   3600
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   1320
         X2              =   3480
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   3000
         X2              =   3000
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFCEC&
      Caption         =   "Table Names"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "AccessBrowser.frx":0004
         Left            =   120
         List            =   "AccessBrowser.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   1320
         X2              =   3480
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   0
         X2              =   3600
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1980
      Left            =   2280
      Pattern         =   "*.mdb"
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdField 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Show Fields"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Browse Table"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELECT ACCESS DATABASE(*.mdb)"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Developed By: Rachit K.
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim xDAO As Database
Dim fileType As String
Dim getPath As String

Private Sub GetMDB(mdbPath As String)
'This is the Function for getting the No. of Tables & their name present in the Database

Dim i, a, filTyp, conString
List1.Clear
On Error GoTo MyErr
Set con = New ADODB.Connection

'Minor Checking if its an Access file
filTyp = IIf(UCase(Right(mdbPath, 3)) = "MDB", "mdb", "dbc")
If filTyp = "mdb" Then
'Access File
    conString = "Provider=Microsoft.Jet.OLEDB.3.51;Password=" & Text2 & ";Persist Security Info=True;Data Source=" & mdbPath
    Set xDAO = OpenDatabase(mdbPath)
End If
con.Open conString

For i = 1 To xDAO.TableDefs.Count
'TableDefs will give me the no. & name of _
tables in DB file
    Set rs = New ADODB.Recordset
    'I am opening this tables here 'coz there r some system tables also which is not suppose to be displayed. _
    it will pop an error if tried 2 open SYS Tables & hence i am tracking the error & filtering just the USER tables
    rs.Open "Select * From " & xDAO.TableDefs(i - 1).Name, con, adOpenStatic, adLockReadOnly
    'Adding Table Names present in the selected Database
    List1.AddItem xDAO.TableDefs(i - 1).Name
    Set rs = Nothing
Next
Exit Sub

MyErr:
If Err.Number = -2147217911 Then
'Get this err when tried 2 open SYS tables
    Exit Sub
Else
    MsgBox "Error Encountered: " & Err.Description
End If
End Sub

Private Sub getFields(xPath As String)
'This is the function for getting the no. & name of the fields in the selected table
On Error GoTo MyErr
Dim conString As String
Dim i As Integer
List2.Clear
If List1.Text <> "" Then
    If fileType = "mdb" Then
       'Access File
        Set xDAO = OpenDatabase(getPath)
        conString = "Provider=Microsoft.Jet.OLEDB.3.51;Password=" & Text2 & ";Persist Security Info=True;Data Source=" & xPath
        Set con = New ADODB.Connection
        con.Open conString
        Set rs = New ADODB.Recordset
        rs.Open "Select * From " & List1.Text, con, adOpenStatic, adLockReadOnly
        For i = 0 To rs.Fields.Count - 1
            List2.AddItem rs(i).Name
        Next
        Set rs = Nothing
    End If
    
End If
Exit Sub

MyErr:

If Err.Number = -2147217911 Then
'Get this err when tried 2 open SYS tables
    Exit Sub
Else
    MsgBox "Error Encountered: " & Err.Description
End If
End Sub

Private Sub Check1_Click()
Text2 = ""

If Check1.Value = 1 Then
    Frame3.Visible = True
Else: Frame3.Visible = False
End If
End Sub

Private Sub cmdBrowse_Click()
If List1.Text <> "" Then
   Call BrowseTable(getPath)
End If
End Sub

Private Sub BrowseTable(mdbPath As String)
'This is the function to Display Data in Grid
On Error GoTo MyErr

Dim i As Integer
Dim conString As String
grid1.Clear
grid1.ClearStructure

If List1.Text <> "" Then
    If fileType = "mdb" Then
       'Access File
        Set xDAO = OpenDatabase(getPath)
        conString = "Provider=Microsoft.Jet.OLEDB.3.51;Password=" & Text2 & ";Persist Security Info=True;Data Source=" & mdbPath
        Set con = New ADODB.Connection
        con.Open conString
        Set rs = New ADODB.Recordset
        rs.Open "Select * From " & List1.Text, con, adOpenStatic, adLockReadOnly
        'To chek if its not empty
        If rs.EOF <> True And rs.BOF <> True Then
           Frame4.Visible = True
           CmdGo.Enabled = False
           File1.Enabled = False
           Drive1.Enabled = False
           Dir1.Enabled = False
           Check1.Enabled = False
           Text2.Enabled = False
           grid1.ColWidth(0) = 200
           Set grid1.DataSource = rs
        Else
            MsgBox "No Records in Table: " & List1.Text
        End If
        Set rs = Nothing
    End If
    
   
End If
Exit Sub

MyErr:
If Err.Number = -2147217911 Then
'Get this err when tried 2 open SYS tables
    Exit Sub
Else
    MsgBox "Error Encountered: " & Err.Description
End If
End Sub

Private Sub cmdField_Click()
Me.MousePointer = vbHourglass
Call getFields(getPath)
Me.MousePointer = vbNormal
End Sub

Private Sub CmdGo_Click()
Me.MousePointer = vbHourglass
getPath = ""
fileType = ""
If File1.FileName <> "" Then
    'Here Checking if double strokes are there
    getPath = IIf(Len(File1.Path) = 3, File1.Path & File1.FileName, File1.Path & "\" & File1.FileName)
    fileType = IIf(UCase(Right(File1.FileName, 3)) = "MDB", "mdb", "dbc")
    Call GetMDB(getPath)
End If
Me.MousePointer = vbNormal
End Sub

Private Sub CmdLink_Click()
   Frame4.Visible = False
   File1.Enabled = True
   Drive1.Enabled = True
   Dir1.Enabled = True
   Check1.Enabled = True
   Text2.Enabled = True
   CmdGo.Enabled = True
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
List1.Clear
List2.Clear
Text2 = ""
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub grid1_EnterCell()
grid1.CellBackColor = &HC0FFC0
End Sub

Private Sub grid1_LeaveCell()
grid1.CellBackColor = &H80000018
End Sub
