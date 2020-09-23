VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmAddUser 
   Caption         =   "Create User"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4947.604
   ScaleMode       =   0  'User
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkScrollContainer cntMain 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11245
      Begin vkUserContolsXP.vkLabel vkLabel6 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Email"
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
      Begin vkUserContolsXP.vkTextBox txtEmail 
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
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
      Begin vkUserContolsXP.vkLabel vkLabel5 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         BackColor       =   16761024
         Caption         =   "Login Creation"
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
      Begin vkUserContolsXP.vkCommand btnCancel 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "Cancel"
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
      Begin vkUserContolsXP.vkCommand btnCreate 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "Create"
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
      Begin VB.ComboBox cmbRole 
         Height          =   315
         ItemData        =   "frmAddUser.frx":4B02
         Left            =   2640
         List            =   "frmAddUser.frx":4B09
         TabIndex        =   5
         Text            =   "ROLE"
         Top             =   2160
         Width           =   2415
      End
      Begin vkUserContolsXP.vkTextBox txtPassword 
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
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
         PassWordChar    =   "*"
         LegendForeColor =   15695701
      End
      Begin vkUserContolsXP.vkTextBox txtUsername 
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
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
      Begin vkUserContolsXP.vkLabel vkLabel4 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Role"
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
      Begin vkUserContolsXP.vkLabel vkLabel3 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password"
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
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Username"
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
      Begin vkUserContolsXP.vkTextBox txtName 
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Name"
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
Attribute VB_Name = "frmAddUser"
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
Dim rsQuery
Dim SQLText As String

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCreate_Click()
    OpenServer
    'Create table
    
    Set rsQuery = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where username='" & txtUsername.Text & "'"
    Set rsQuery = conn.Execute(SQLText)
         
    If Not rsQuery.EOF Then
        MsgBox "Please use other username."
        txtUsername.Text = ""
    ElseIf txtUsername.Text = "" And txtName.Text = "" And txtPassword.Text = "" And txtEmail.Text = "" And cmbRole.Text = "ROLE" Then
        MsgBox "Please insert the empty text box."
    Else
        conn.Execute "INSERT INTO staff " _
        & "(name, username,password,role,email) VALUES " _
        & "('" & txtName.Text & "','" & txtUsername.Text & "',SHA1('" & txtPassword.Text & "'),'" & cmbRole.Text & "','" & txtEmail.Text & "' )", , adExecuteNoRecords
        
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    OpenFileConfig
    OpenServer ' Open without ODBC in Control Panel
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
