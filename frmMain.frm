VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmMain 
   Caption         =   "Login"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      Begin VB.TextBox txtStaffID 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtRole 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin vkUserContolsXP.vkCommand btnRegister 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "REGISTER USER"
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
      Begin vkUserContolsXP.vkCommand btnLogin 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "LOGIN"
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
      Begin vkUserContolsXP.vkTextBox txtPassword 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   840
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
         PassWordChar    =   "*"
         LegendForeColor =   15695701
      End
      Begin vkUserContolsXP.vkTextBox txtUsername 
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   480
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
      Begin vkUserContolsXP.vkLabel vkLabel2 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password :"
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
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Username :"
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
Attribute VB_Name = "frmMain"
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

Private Sub btnLogin_Click()
    Dim staff_username As String
    Dim staff_password As String
    Dim staff_role As String
    Dim staff_id As Long
    
    OpenServer
    'Create table
    
    Set rsQuery = CreateObject("ADODB.Recordset")
    SQLText = "SELECT StaffID,username,password,role FROM staff where username='" & txtUsername.Text & "' and password=SHA1('" & txtPassword.Text & "')"
    Set rsQuery = conn.Execute(SQLText)
         
    If Not rsQuery.EOF Then
        staff_role = rsQuery("role")
        staff_id = rsQuery("StaffID")
        txtRole.Text = staff_role
        txtStaffID.Text = staff_id
        frmMain2.Show
        Unload Me
    Else
        MsgBox "Wrong Username or Password please try again."
        txtUsername.Text = ""
        txtPassword.Text = ""
    End If
    
    
End Sub

Private Sub btnRegister_Click()
    frmAddUser.Show
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
