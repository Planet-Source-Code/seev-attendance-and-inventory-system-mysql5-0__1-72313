VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmSeachReceipt 
   Caption         =   "Search Receipt"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8850
   Icon            =   "frmSeachReceipt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      Height          =   5655
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9975
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin vkUserContolsXP.vkCommand btnSearch 
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "SEARCH"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkTextBox txtSession 
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendForeColor =   15695701
      End
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Receipt No."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSeachReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db_name As String
Dim db_server As String
Dim db_port As String
Dim db_user As String
Dim db_pass As String
Dim db_driver As String
Dim constr As String
Dim conn As ADODB.Connection
Dim rsQuery1
Dim rsQuery2
Dim rsQuery3
Dim rsQuery4
Dim SQLText As String
Dim staff_id As Long
Dim staff_id2 As Long
Dim status_ As String
Dim Rx As Long
Dim RS_Staff As ADODB.Recordset

Dim name2 As String
Dim username As String
Dim email As String

Dim address As String
Dim city As String
Dim State As String
Dim zipcode As String

Dim tel_home As String
Dim tel_mobile As String
Dim position As String
Dim commission As String
Dim salary_scale As Long
Dim hourly_rate As Long

Private Sub btnSearch_Click()
    Dim rs2
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where session='" & txtSession.Text & "'"
    Set rsQuery1 = conn.Execute(SQLText)
    
    If Not rsQuery1.EOF Then
        Set rs2 = New ADODB.Recordset
        rs2.CursorLocation = adUseClient
        rs2.CursorType = adOpenStatic
        rs2.LockType = adLockReadOnly
        rs2.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, total_price as Price FROM item_sale where session='" & txtSession.Text & "'", conn
        Set DataGrid1.DataSource = rs2
    Else
        MsgBox ("No Record!!!")
    End If
End Sub

Private Sub Form_Load()
    OpenFileConfig
    OpenServer ' Open without ODBC in Control Panel
End Sub

Private Sub OpenFileConfig()
  On Error GoTo eexit
    
   Dim path As String
   Dim aRecord As String
   Dim EqualTo As Integer
   
  Open path & "Config.ebu" For Input As #2
  
   Do Until (EOF(2) = True)
     Input #2, aRecord
     EqualTo = InStr(aRecord, "=")
    
    Select Case Left(aRecord, EqualTo - 1)
     
    'Config DB
    Case "db_name"
        db_name = Mid(aRecord, EqualTo + 1)
    Case "db_server"
        db_server = Mid(aRecord, EqualTo + 1)
    Case "db_port"
        db_port = Mid(aRecord, EqualTo + 1)
    Case "db_user"
        db_user = Mid(aRecord, EqualTo + 1)
    Case "db_pass"
        db_pass = Mid(aRecord, EqualTo + 1)
    Case "db_driver"
        db_driver = Mid(aRecord, EqualTo + 1)

    End Select
  Loop
eexit:
  Close 2
End Sub

Private Sub OpenServer() 'Connect MySQL Server Without ODBC setup
On Error GoTo DBerror
    
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=" & db_user & ";Extended Properties=" & Chr$(34) & "DRIVER={" & db_driver & "};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    
    Exit Sub

DBerror:
    frmSetupDB.Show
    Unload Me
End Sub


