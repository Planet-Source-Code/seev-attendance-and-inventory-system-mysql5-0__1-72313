VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmSetupDB 
   Caption         =   "Setup Database"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11880
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4260
         Caption         =   "Processing"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin vkUserContolsXP.vkListBox List1 
            Height          =   1275
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2249
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Sorted          =   0
         End
      End
      Begin vkUserContolsXP.vkCommand btnCancel 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
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
      Begin vkUserContolsXP.vkCommand btnCreateDB 
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Create Database"
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
      Begin VB.TextBox txtDBPASSWORD 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cmbDriver 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "main.frx":4B02
         Left            =   2640
         List            =   "main.frx":4B0F
         TabIndex        =   8
         Text            =   "MYSQL Only"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox cmbPort 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "main.frx":4B49
         Left            =   2640
         List            =   "main.frx":4B59
         TabIndex        =   7
         Text            =   "Your Port"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtDBUSERNAME 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtDBSERVER 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtDBNAME 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "sales"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MySQL Password :"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "MySQL Driver :"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Port : "
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MySQL Username :"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name :"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSetupDB"
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
Private Sub btnCreateDB_Click()
    db_name = txtDBNAME.Text
    db_server = txtDBSERVER.Text
    db_port = cmbPort.Text
    db_user = txtDBUSERNAME.Text
    db_pass = txtDBPASSWORD.Text
    db_driver = cmbDriver.Text
    
    CreateDB
    CreateTable
    SaveFile
    
    MsgBox "Your username : asministrator and your password is : admin1234. Please change when you logon."
    
    Sleep (1000)
    frmMain.Show
    Unload Me
End Sub

Private Sub btnCancel_Click()
    End
End Sub

Private Sub CreateDB()
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=" & db_user & ";Extended Properties=" & Chr$(34) & "DRIVER={" & db_driver & "};DESC=;SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    conn.Execute "DROP DATABASE IF EXISTS " & db_name
    conn.Execute "Create Database " & Trim$(db_name), , adExecuteNoRecords
    conn.Execute "SHOW DATABASES"
    
    WriteList "Creating Database....." & db_name & " Success"
        
End Sub
Private Sub CreateTable()
    OpenServer
    'Create table
    conn.Execute "DROP TABLE IF EXISTS staff"

    WriteList "Creating Table....."
    
    conn.Execute "CREATE TABLE staff " _
    & "(staffID BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "name VARCHAR (50)," _
    & "username VARCHAR (50)," _
    & "email VARCHAR (100)," _
    & "password VARCHAR (200)," _
    & "role VARCHAR (20)," _
    & "reg_date TIMESTAMP," _
    & "PRIMARY KEY(staffID), UNIQUE(staffID), INDEX(staffID)) TYPE = MyISAM;", , adExecuteNoRecords
    
    conn.Execute "INSERT INTO staff(username,password,role,email,name) VALUES ('administrator',SHA1('admin1234'),'Admin','administrator@localhost','Administrator')", , adExecuteNoRecords
    
    WriteList "Create <tbluser> Success"
    
    conn.Execute "DROP TABLE IF EXISTS staff_address"
    
    conn.Execute "CREATE TABLE staff_address " _
    & "(id BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "address VARCHAR (200)," _
    & "city VARCHAR (50)," _
    & "state VARCHAR (50)," _
    & "zipcode VARCHAR (50)," _
    & "staffID BIGINT (17)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(id), UNIQUE(id), INDEX(id)) TYPE = MyISAM;", , adExecuteNoRecords
    
    WriteList "Create <staff_address> Success"
    
    conn.Execute "DROP TABLE IF EXISTS staff_details"
    
    conn.Execute "CREATE TABLE staff_details " _
    & "(id BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "tel_home VARCHAR (20)," _
    & "tel_mobile VARCHAR (20)," _
    & "position VARCHAR (50)," _
    & "commission VARCHAR (50)," _
    & "salary_scale BIGINT (10)," _
    & "hourly_rate BIGINT (10)," _
    & "status VARCHAR (50)," _
    & "staffID BIGINT (17)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(id), UNIQUE(id), INDEX(id)) TYPE = MyISAM;", , adExecuteNoRecords

    WriteList "Create <staff_details> Success"
    
    conn.Execute "DROP TABLE IF EXISTS staff_attendance"
    
    conn.Execute "CREATE TABLE staff_attendance " _
    & "(id BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "time_start DATETIME NOT NULL," _
    & "time_finish DATETIME DEFAULT '1990-01-01 01:01:01' NOT NULL," _
    & "hours_working INT (10) DEFAULT '1' NOT NULL," _
    & "staffID BIGINT (17)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(id), UNIQUE(id), INDEX(id)) TYPE = MyISAM;", , adExecuteNoRecords

    WriteList "Create <staff_attendance> Success"
    
    conn.Execute "DROP TABLE IF EXISTS sales_reference"
    
    conn.Execute "CREATE TABLE sales_reference " _
    & "(salesID BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "sales_amount BIGINT (50)," _
    & "staffID BIGINT (17)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(salesID), UNIQUE(salesID), INDEX(salesID)) TYPE = MyISAM;", , adExecuteNoRecords

    WriteList "Create <sales_reference> Success"
    
    conn.Execute "DROP TABLE IF EXISTS item_sale"
    
    conn.Execute "CREATE TABLE item_sale " _
    & "(id BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "quantity_sold BIGINT (50)," _
    & "salesID BIGINT (17)," _
    & "total_price DOUBLE DEFAULT NULL," _
    & "price DOUBLE DEFAULT NULL," _
    & "item_desc VARCHAR (200)," _
    & "staffID BIGINT (17)," _
    & "session VARCHAR(50)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(id), UNIQUE(id), INDEX(id)) TYPE = MyISAM;", , adExecuteNoRecords

    WriteList "Create <item_sale> Success"
    
    conn.Execute "DROP TABLE IF EXISTS inventory"
    
    conn.Execute "CREATE TABLE inventory " _
    & "(id BIGINT (17) UNSIGNED NOT NULL AUTO_INCREMENT," _
    & "item_desc VARCHAR (200)," _
    & "price DOUBLE DEFAULT NULL," _
    & "quantity_instock BIGINT (50)," _
    & "minimum_order_quantity BIGINT (50)," _
    & "type_of_item VARCHAR (200)," _
    & "staffID BIGINT (17)," _
    & "submit_date TIMESTAMP," _
    & "PRIMARY KEY(id), UNIQUE(id), INDEX(id)) TYPE = MyISAM;", , adExecuteNoRecords

    WriteList "Create <inventory> Success"
    
    Exit Sub
End Sub
Private Sub OpenServer() 'Connect MySQL Server Without ODBC setup

On Error GoTo DBerror
    
    constr = "Provider=MSDASQL.1;Password=;Persist Security Info=True;User ID=;Extended Properties=" & Chr$(34) & "DRIVER={" & db_driver & "};DESC=;DATABASE=" & db_name & ";SERVER=" & db_server & ";UID=" & db_user & ";PASSWORD=" & db_pass & ";PORT=" & db_port & ";OPTION=16387;STMT=;" & Chr$(34)
    Set conn = New ADODB.Connection
    conn.Open constr
    WriteList "Connect to DB : " & db_name & " on server : " & db_server
    
    Exit Sub

DBerror:
    WriteList "Failed to connect on DB : " & db_name & " on server : " & db_server
    Sleep (1000)
    End
End Sub

Private Sub SaveFile()
    Open ConfigPath & "Config.ebu" For Output As #1
      
    'Config DB
    Write #1, "db_name=" & db_name
    Write #1, "db_server=" & db_server
    Write #1, "db_port=" & db_port
    Write #1, "db_user=" & db_user
    Write #1, "db_pass=" & db_pass
    Write #1, "db_driver=" & db_driver
    
    WriteList "Save configuration to : " & ConfigPath & "Config.ebu"
eexit:
    Close 1
End Sub
Public Sub WriteList(ByVal aText As String)
     On Error Resume Next
    
     If List1.ListCount > 200 Then
        List1.RemoveItem 0
    End If
    List1.AddItem aText
    List1.ListIndex = List1.ListCount - 1

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

Private Sub Form_Load()
    OpenFileConfig
    frmSetupDB.Show
    txtDBNAME.Text = db_name
    txtDBSERVER.Text = db_server
    cmbPort.Text = db_port
    txtDBUSERNAME.Text = db_user
    txtDBPASSWORD.Text = db_pass
    cmbDriver.Text = db_driver
End Sub
