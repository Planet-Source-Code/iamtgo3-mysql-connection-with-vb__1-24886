VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Bindings        =   "Form1.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   120
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=myIPDG3"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "myIPDG3"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "lucinda3"
      RecordSource    =   "SELECT * FROM feedback"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   7695
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7725
      Width           =   9375
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   4440
         Picture         =   "Form1.frx":0015
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   4785
         Picture         =   "Form1.frx":0357
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   8520
         Picture         =   "Form1.frx":0699
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   8865
         Picture         =   "Form1.frx":09DB
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5130
         TabIndex        =   9
         Top             =   120
         Width           =   3360
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Index           =   3
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "Form1.frx":0D1D
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Subject :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Message :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Email Address :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAbort 
         Caption         =   "Abort"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditAddNewRecord 
         Caption         =   "Add New Record"
      End
      Begin VB.Menu mnuEditEditRecord 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu mnuEditDeleteRecord 
         Caption         =   "Delete Record"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program requires that you have mySQL loaded somewhere you can connect to it.
' You must also have myODBC loaded on your machine. This program uses Windows ODBC
' to connect to your mySQL database. I just put a bunch of controls on a form to play
' around with it and thought others might be interested in looking at it.

Option Explicit

Dim mvBookmark1 As Variant
Dim mvBookmark2 As Variant

Dim mstQuery As String

Public con As New ADODB.Connection
Public cmd As New ADODB.Command
Public rs As New ADODB.Recordset

Public myCon As New ADODB.Connection
Public myCmd As New ADODB.Command
Public myRS As New ADODB.Recordset

Private Sub SizeColumns(ByVal flx As MSHFlexGrid)
    Dim max_wid As Single
    Dim wid As Single
    Dim max_row As Integer
    Dim r As Integer
    Dim c As Integer
    
    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
       max_wid = 0
       For r = 0 To max_row
          wid = TextWidth(flx.TextMatrix(r, c))
          If max_wid < wid Then max_wid = wid
       Next r
       flx.ColWidth(c) = max_wid + 240
    Next c
End Sub

Private Sub cmdFirst_Click()

    rs.MoveFirst
    DisplayData

End Sub

Private Sub cmdLast_Click()

    rs.MoveLast
    DisplayData

End Sub

Private Sub cmdNext_Click()
  
    If Not rs.EOF Then rs.MoveNext

    If rs.EOF And rs.RecordCount > 0 Then
      Beep
      'moved off the end so go back
      rs.MoveLast
    End If

    DisplayData

End Sub

Private Sub cmdPrevious_Click()
  
    If Not rs.BOF Then rs.MovePrevious
    
    If rs.BOF And rs.RecordCount > 0 Then
      Beep
      'moved off the end so go back
      rs.MoveFirst
    End If
    
    DisplayData

End Sub

Private Sub Form_Load()

    On Error GoTo Error
    
    con.Open "DSN=myDNS"

    With cmd
      Set .ActiveConnection = con
      .CommandType = adCmdTable
      .CommandText = "feedback"
    End With

    With rs
      .LockType = adLockPessimistic
      .CursorType = adOpenKeyset
      .CursorLocation = adUseClient
      .Open cmd
    End With

    mstQuery = "SELECT * FROM feedback"
    
    
    myCon.Open "DSN=myDNS"

    With myCmd
      Set .ActiveConnection = myCon
      .CommandType = adCmdText
      .CommandText = mstQuery
    End With

    With myRS
      .LockType = adLockPessimistic
      .CursorType = adOpenKeyset
      .CursorLocation = adUseClient
      .Open myCmd
    End With

    Set Adodc1.Recordset = myRS
    Adodc1.Refresh
    myRS.Close

    SizeColumns MSHFlexGrid1
    MSHFlexGrid1.Refresh

    SizeColumns MSHFlexGrid2
    MSHFlexGrid2.Refresh

    rs.MoveFirst

    SetRecordNumber

    DisplayData
    
    On Error GoTo 0
    
Form_Load_Exit:
    Exit Sub
    
Error:
    MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, "Error in [Form_Load]"

End Sub

Private Sub SetRecordNumber()

    lblStatus.Caption = "Record: " & CStr(rs.AbsolutePosition) & " of " & CStr(rs.RecordCount)
    
End Sub

Private Sub mnuEditAddNewRecord_Click()
Dim iCount As Integer

    On Error GoTo Error
    
'    miAddEditIdentical = 1
    
    For iCount = 0 To 3
      Text1(iCount).Text = ""
    Next

    rs.AddNew
        
'    SetButtons False

    Text1(0).SetFocus
'
'    giAddToFile = 1
    
    On Error GoTo 0

mnuEditAddNewRecord_Exit:
    Exit Sub

Error:
    MsgBox Err.Description, vbExclamation, "Error in [mnuEditAddNewRecord]"
'    SetButtons True

End Sub

Private Sub mnuEditDeleteRecord_Click()
    
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Are You Sure You Want to Delete Current Record ?"   ' Define message.
    Style = vbYesNo + vbCritical ' Define buttons.
    Title = "Confirm Deletion"  ' Define title.
    
    If rs.RecordCount > 1 Then
      Response = MsgBox(Msg, Style, Title, Help, Ctxt)
      If Response = vbYes Then
        With rs
          .Delete
          .MoveNext
          If .EOF Then
            .MovePrevious
            If .BOF Then
              MsgBox "Your Database is Empty", vbInformation, "No Records"
            End If
          End If
        End With
      End If
    Else
      MsgBox "You only have one record in the database you must Add a record before you delete this one...", vbInformation, "Add Record First"
    End If
    
    DisplayData

End Sub

Private Sub mnuEditEditRecord_Click()
   
    On Error GoTo Error
    
'    rs.Edit
    
'    SetButtons False

    Text1(0).SetFocus
    
'    giAddToFile = 3
    
    On Error GoTo 0

mnuEditEditRecord_Exit:
    Exit Sub

Error:
    MsgBox Err.Description, vbExclamation, "Error in [mnuEditEditRecord]"
'    SetButtons True

End Sub

Private Sub mnuFileAbort_Click()

    On Error GoTo Error
    
    rs.CancelUpdate
    
    cmdFirst_Click
    
'    SetButtons True
    
    On Error GoTo 0

mnuFileAbort_Exit:
    Exit Sub

Error:
    MsgBox Err.Description, vbExclamation, "Error in [mnuFileAbort]"
'    SetButtons True


End Sub

Private Sub mnuFileExit_Click()

    rs.Close

    myCon.Close
    con.Close

    End

End Sub

Private Sub DisplayData()

    On Error GoTo Error
    
'    If rs.BOF Then rs.Bookmark = mvBookmark1
'    If rs.EOF Then rs.Bookmark = mvBookmark1
    
    If IsNull(rs("username")) Then
      Text1(0).Text = ""
    Else
      Text1(0).Text = rs("username")
    End If
      
    If IsNull(rs("emailaddress")) Then
      Text1(1).Text = ""
    Else
      Text1(1).Text = rs("emailaddress")
    End If
      
    If IsNull(rs("subject")) Then
      Text1(2).Text = ""
    Else
      Text1(2).Text = rs("subject")
    End If
      
    If IsNull(rs("message")) Then
      Text1(3).Text = ""
    Else
      Text1(3).Text = rs("message")
    End If
      
    SetRecordNumber
    
    On Error GoTo 0

DisplayData_Exit:
    Exit Sub
    
Error:
    MsgBox Err.Description, vbExclamation, "Error in [DisplayData]"
    
End Sub

Private Sub mnuFileSave_Click()
    
'    On Error GoTo Error
    
'    StripOutApostrophes
    
    rs("username") = Text1(0).Text
    rs("emailaddress") = Text1(1).Text
    rs("subject") = Text1(2).Text
    rs("message") = Text1(3).Text
    rs.Update
    
    SetRecordNumber
    
'    SetButtons True
   
    On Error GoTo 0

mnuFileSave_Exit:
    Exit Sub

Error:
    MsgBox Err.Description, vbExclamation, "Error in [mnuFileSave]"
'    SetButtons True
    mnuFileAbort_Click

End Sub
