VERSION 5.00
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmUpInventory 
   Caption         =   "Update Inventory"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6225
   Icon            =   "frmUpInventory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkCommand btnFind 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Find"
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
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Inventory ID"
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
   Begin vkUserContolsXP.vkTextBox txtInvID 
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
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
   Begin vkUserContolsXP.vkFrame vkFrame5 
      Height          =   3735
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      Caption         =   "Adding Inventory"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleGradient   =   2
      Begin vkUserContolsXP.vkCommand btnUpInventory 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   3120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "UPDATE INVENTORY"
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
      Begin vkUserContolsXP.vkTextBox txtTypeOfItem 
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   2520
         Width           =   2535
         _ExtentX        =   4471
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
      Begin vkUserContolsXP.vkTextBox txtMinOrder 
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
         _ExtentX        =   4471
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
      Begin vkUserContolsXP.vkTextBox txtQIS 
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
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
      Begin vkUserContolsXP.vkTextBox txtPrice 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
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
      Begin vkUserContolsXP.vkTextBox txtItemDesc 
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin vkUserContolsXP.vkLabel vkLabel54 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Type of Item"
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
      Begin vkUserContolsXP.vkLabel vkLabel55 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Minimum Order Stock"
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
      Begin vkUserContolsXP.vkLabel vkLabel56 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Quantity in Stock"
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
      Begin vkUserContolsXP.vkLabel vkLabel57 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Price"
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
      Begin vkUserContolsXP.vkLabel vkLabel58 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Item Description"
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
Attribute VB_Name = "frmUpInventory"
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

Private Sub btnFind_Click()
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM inventory where id='" & txtInvID.Text & "'"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        txtItemDesc.Text = rsQuery1("item_desc")
        txtPrice.Text = rsQuery1("price")
        txtMinOrder.Text = rsQuery1("minimum_order_quantity")
        txtTypeOfItem.Text = rsQuery1("type_of_item")
        txtQIS.Text = rsQuery1("quantity_instock")
    Else
        MsgBox ("No Record!!!")
    End If
End Sub

Private Sub btnUpInventory_Click()
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM inventory"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If IsNumeric(txtPrice.Text) = True And IsNumeric(txtMinOrder.Text) And IsNumeric(txtQIS.Text) Then
        If Not rsQuery1.EOF Then
            conn.Execute "UPDATE inventory SET" _
            & " item_desc='" & txtItemDesc.Text & "' " _
            & " ,price='" & txtPrice.Text & "' " _
            & " ,minimum_order_quantity='" & txtMinOrder.Text & "' " _
            & " ,type_of_item='" & txtTypeOfItem.Text & "' " _
            & " ,quantity_instock='" & txtQIS.Text & "' WHERE id='" & txtInvID.Text & "' ", , adExecuteNoRecords
            MsgBox "Inventory is updated"
        Else
            MsgBox "Error"
        End If
    Else
        MsgBox ("Please put numeric value only for Price or Quantity In Stock or Minimum Order Quantity")
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
