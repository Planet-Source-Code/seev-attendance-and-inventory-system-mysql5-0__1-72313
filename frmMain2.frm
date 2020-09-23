VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{EAFDAFBF-1D88-41DD-B117-60ECBC4B8441}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain2 
   Caption         =   "Sales System V1.0"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13425
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleMode       =   0  'User
   ScaleWidth      =   2294.872
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer1 
      DragMode        =   1  'Automatic
      Height          =   7335
      Left            =   113
      TabIndex        =   9
      Top             =   1320
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   12938
      BackColor2      =   16761024
      Begin vkUserContolsXP.vkLabel vkLabel1 
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   450
         BorderStyle     =   1
         BackColor       =   14737632
         Caption         =   "    Action Container"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabDlg.SSTab tabSales 
         Height          =   6735
         Left            =   120
         TabIndex        =   178
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Add Item"
         TabPicture(0)   =   "frmMain2.frx":4B02
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vkFrame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "btnPrevItemAdmin"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnNextItemAdmin"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnDelItemAdmin"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblTotalPriceAdmin"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "vkLabel59"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtIIDAdmin"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CommonDialog1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "btnPrintReceipt"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtSessionAdmin"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "vkLabel60"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "tbSearchRAdmin"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "View Report"
         TabPicture(1)   =   "frmMain2.frx":4B1E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmbYrSales"
         Tab(1).Control(1)=   "cmbMthSales"
         Tab(1).Control(2)=   "cdExportSales"
         Tab(1).Control(3)=   "btnYrSales"
         Tab(1).Control(4)=   "btnMthSales"
         Tab(1).Control(5)=   "DataGrid7"
         Tab(1).Control(6)=   "btnShowReportSales"
         Tab(1).Control(7)=   "btnExportXLSSales"
         Tab(1).ControlCount=   8
         Begin VB.ComboBox cmbYrSales 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4B3A
            Left            =   -73200
            List            =   "frmMain2.frx":4B80
            TabIndex        =   213
            Text            =   "Year"
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox cmbMthSales 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4C08
            Left            =   -74880
            List            =   "frmMain2.frx":4C30
            TabIndex        =   212
            Text            =   "Month"
            Top             =   480
            Width           =   1575
         End
         Begin vkUserContolsXP.vkCommand tbSearchRAdmin 
            Height          =   375
            Left            =   10680
            TabIndex        =   202
            Top             =   5520
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Search Receipt"
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
         Begin vkUserContolsXP.vkLabel vkLabel60 
            Height          =   255
            Left            =   10680
            TabIndex        =   200
            Top             =   3240
            Width           =   2055
            _ExtentX        =   3625
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
         Begin VB.TextBox txtSessionAdmin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10680
            TabIndex        =   195
            Top             =   3600
            Width           =   2055
         End
         Begin vkUserContolsXP.vkCommand btnPrintReceipt 
            Height          =   375
            Left            =   10680
            TabIndex        =   194
            Top             =   5040
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Print Receipt"
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6600
            Top             =   6000
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtIIDAdmin 
            Height          =   285
            Left            =   10680
            TabIndex        =   181
            Top             =   2880
            Visible         =   0   'False
            Width           =   2055
         End
         Begin vkUserContolsXP.vkLabel vkLabel59 
            Height          =   255
            Left            =   120
            TabIndex        =   179
            Top             =   6000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Total Price :"
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
         Begin vkUserContolsXP.vkLabel lblTotalPriceAdmin 
            Height          =   255
            Left            =   2160
            TabIndex        =   180
            Top             =   6000
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            BackColor       =   14737632
            Caption         =   "RM 0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkCommand btnDelItemAdmin 
            Height          =   495
            Left            =   10680
            TabIndex        =   182
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BackColor1      =   8421631
            BackColor2      =   14345190
            BackColorPushed1=   12632319
            BackColorPushed2=   14345442
            BackGradient    =   0
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   0
         End
         Begin vkUserContolsXP.vkCommand btnNextItemAdmin 
            Height          =   495
            Left            =   11760
            TabIndex        =   183
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Next"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin vkUserContolsXP.vkCommand btnPrevItemAdmin 
            Height          =   495
            Left            =   10680
            TabIndex        =   184
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Previous"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin MSDataGridLib.DataGrid DataGrid5 
            Height          =   4335
            Left            =   120
            TabIndex        =   185
            Top             =   1560
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7646
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
         Begin vkUserContolsXP.vkFrame vkFrame6 
            Height          =   1215
            Left            =   120
            TabIndex        =   186
            Top             =   240
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   2143
            Caption         =   "Adding New Item"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TitleColor2     =   16744576
            TitleGradient   =   2
            Begin vkUserContolsXP.vkLabel vkLabel63 
               Height          =   255
               Left            =   1320
               TabIndex        =   193
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkLabel vkLabel62 
               Height          =   255
               Left            =   4200
               TabIndex        =   192
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Quantity"
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
            Begin vkUserContolsXP.vkLabel vkLabel61 
               Height          =   255
               Left            =   7080
               TabIndex        =   191
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkTextBox txtItemIDAdmin 
               Height          =   255
               Left            =   1320
               TabIndex        =   190
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkTextBox txtQuantityAdmin 
               Height          =   255
               Left            =   4200
               TabIndex        =   189
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkTextBox txtItemDescAddAdmin 
               Height          =   255
               Left            =   7080
               TabIndex        =   188
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnAddItemAdmin 
               Height          =   615
               Left            =   11280
               TabIndex        =   187
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1085
               Caption         =   "ADD"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin MSComDlg.CommonDialog cdExportSales 
            Left            =   -69960
            Top             =   6120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin vkUserContolsXP.vkCommand btnYrSales 
            Height          =   375
            Left            =   -73200
            TabIndex        =   214
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Yearly"
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
         Begin vkUserContolsXP.vkCommand btnMthSales 
            Height          =   375
            Left            =   -74880
            TabIndex        =   215
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Monthly"
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
         Begin MSDataGridLib.DataGrid DataGrid7 
            Height          =   4575
            Left            =   -75000
            TabIndex        =   216
            Top             =   1080
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   8070
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   6
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
         Begin vkUserContolsXP.vkCommand btnShowReportSales 
            Height          =   375
            Left            =   -71400
            TabIndex        =   217
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Show Report"
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
         Begin vkUserContolsXP.vkCommand btnExportXLSSales 
            Height          =   375
            Left            =   -68880
            TabIndex        =   218
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Export to Excel"
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
      Begin TabDlg.SSTab tabSalesStaff 
         Height          =   6735
         Left            =   120
         TabIndex        =   126
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Add Item"
         TabPicture(0)   =   "frmMain2.frx":4C64
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CommonDialog2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnPrintReceiptStaff"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vkFrame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnPrevStaff"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnNextStaff"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "btnDeleteSalesStaff"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtItemIDStaff"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblTotPrice"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "vkLabel48"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtSessionStaff"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "vkLabel64"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "btnSearchRS"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         Begin vkUserContolsXP.vkCommand btnSearchRS 
            Height          =   375
            Left            =   10680
            TabIndex        =   203
            Top             =   5520
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Search Receipt"
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
         Begin vkUserContolsXP.vkLabel vkLabel64 
            Height          =   255
            Left            =   10680
            TabIndex        =   201
            Top             =   3240
            Width           =   2055
            _ExtentX        =   3625
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
         Begin VB.TextBox txtSessionStaff 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10680
            TabIndex        =   196
            Top             =   3600
            Width           =   2055
         End
         Begin vkUserContolsXP.vkLabel vkLabel48 
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   6000
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Total Price :"
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
         Begin vkUserContolsXP.vkLabel lblTotPrice 
            Height          =   255
            Left            =   2160
            TabIndex        =   140
            Top             =   6000
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            BackColor       =   14737632
            Caption         =   "RM 0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin VB.TextBox txtItemIDStaff 
            Height          =   285
            Left            =   10680
            TabIndex        =   139
            Top             =   2880
            Visible         =   0   'False
            Width           =   2055
         End
         Begin vkUserContolsXP.vkCommand btnDeleteSalesStaff 
            Height          =   495
            Left            =   10680
            TabIndex        =   138
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BackColor1      =   8421631
            BackColor2      =   14345190
            BackColorPushed1=   12632319
            BackColorPushed2=   14345442
            BackGradient    =   0
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   0
         End
         Begin vkUserContolsXP.vkCommand btnNextStaff 
            Height          =   495
            Left            =   11760
            TabIndex        =   137
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Next"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin vkUserContolsXP.vkCommand btnPrevStaff 
            Height          =   495
            Left            =   10680
            TabIndex        =   136
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Previous"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   4335
            Left            =   120
            TabIndex        =   135
            Top             =   1560
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7646
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
         Begin vkUserContolsXP.vkFrame vkFrame3 
            Height          =   1215
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   2143
            Caption         =   "Adding New Item"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TitleColor2     =   16744576
            TitleGradient   =   2
            Begin vkUserContolsXP.vkCommand btnAddItemStaff 
               Height          =   615
               Left            =   11280
               TabIndex        =   134
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1085
               Caption         =   "ADD"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin vkUserContolsXP.vkTextBox txtItemDescAddS 
               Height          =   255
               Left            =   7080
               TabIndex        =   133
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtQuantityAddS 
               Height          =   255
               Left            =   4200
               TabIndex        =   132
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkTextBox txtItemIDAddS 
               Height          =   255
               Left            =   1320
               TabIndex        =   131
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkLabel vkLabel47 
               Height          =   255
               Left            =   7080
               TabIndex        =   130
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
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
            Begin vkUserContolsXP.vkLabel vkLabel46 
               Height          =   255
               Left            =   4200
               TabIndex        =   129
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Quantity"
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
            Begin vkUserContolsXP.vkLabel vkLabel45 
               Height          =   255
               Left            =   1320
               TabIndex        =   128
               Top             =   360
               Width           =   2655
               _ExtentX        =   4683
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
         End
         Begin vkUserContolsXP.vkCommand btnPrintReceiptStaff 
            Height          =   375
            Left            =   10680
            TabIndex        =   197
            Top             =   5040
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Print Receipt"
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
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   6600
            Top             =   6000
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin TabDlg.SSTab tabStaff 
         Height          =   6735
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Update Staff Info"
         TabPicture(0)   =   "frmMain2.frx":4C80
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmUserProfile"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "vkLabel24"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtUsernameFind"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnFind"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Add New Staff"
         TabPicture(1)   =   "frmMain2.frx":4C9C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vkFrame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Staff Report"
         TabPicture(2)   =   "frmMain2.frx":4CB8
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cdExportStaff"
         Tab(2).Control(1)=   "lblTotalSalary"
         Tab(2).Control(2)=   "vkLabel44"
         Tab(2).Control(3)=   "cmbYear"
         Tab(2).Control(4)=   "cmbMonth"
         Tab(2).Control(5)=   "btnYearly"
         Tab(2).Control(6)=   "btnMonthly"
         Tab(2).Control(7)=   "DataGrid1"
         Tab(2).Control(8)=   "btnRetrieve"
         Tab(2).Control(9)=   "btnExportCSVStaff"
         Tab(2).ControlCount=   10
         Begin MSComDlg.CommonDialog cdExportStaff 
            Left            =   -70320
            Top             =   6000
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin vkUserContolsXP.vkLabel lblTotalSalary 
            Height          =   255
            Left            =   -73080
            TabIndex        =   125
            Top             =   6000
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            Caption         =   "RM 0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkLabel vkLabel44 
            Height          =   255
            Left            =   -74880
            TabIndex        =   124
            Top             =   6000
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Total Salary"
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
         Begin VB.ComboBox cmbYear 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4CD4
            Left            =   -73200
            List            =   "frmMain2.frx":4D1A
            TabIndex        =   123
            Text            =   "Year"
            Top             =   780
            Width           =   1695
         End
         Begin VB.ComboBox cmbMonth 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4DA2
            Left            =   -74880
            List            =   "frmMain2.frx":4DCA
            TabIndex        =   122
            Text            =   "Month"
            Top             =   780
            Width           =   1575
         End
         Begin vkUserContolsXP.vkCommand btnYearly 
            Height          =   375
            Left            =   -73200
            TabIndex        =   121
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Yearly"
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
         Begin vkUserContolsXP.vkCommand btnMonthly 
            Height          =   375
            Left            =   -74880
            TabIndex        =   120
            Top             =   300
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Monthly"
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
         Begin vkUserContolsXP.vkCommand btnFind 
            Height          =   255
            Left            =   4200
            TabIndex        =   42
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
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
         Begin vkUserContolsXP.vkTextBox txtUsernameFind 
            Height          =   255
            Left            =   1560
            TabIndex        =   43
            Top             =   240
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
         Begin vkUserContolsXP.vkLabel vkLabel24 
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Search Username"
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
         Begin vkUserContolsXP.vkFrame frmUserProfile 
            Height          =   5655
            Left            =   120
            TabIndex        =   45
            Top             =   660
            Visible         =   0   'False
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   9975
            Caption         =   "Edit Staff Profile"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox cmbStatusU 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmMain2.frx":4DFE
               Left            =   9360
               List            =   "frmMain2.frx":4E08
               TabIndex        =   119
               Text            =   "STATUS"
               Top             =   2400
               Width           =   3015
            End
            Begin vkUserContolsXP.vkLabel vkLabel43 
               Height          =   255
               Left            =   7080
               TabIndex        =   118
               Top             =   2400
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Status"
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
            Begin vkUserContolsXP.vkLabel vkLabel13 
               Height          =   255
               Left            =   240
               TabIndex        =   78
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel14 
               Height          =   255
               Left            =   240
               TabIndex        =   77
               Top             =   960
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel15 
               Height          =   255
               Left            =   240
               TabIndex        =   76
               Top             =   1320
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkTextBox txtNameAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   75
               Top             =   600
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtUsernameAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   74
               Top             =   960
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtEmailAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   73
               Top             =   1320
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnUp1Admin 
               Height          =   375
               Left            =   5400
               TabIndex        =   72
               Top             =   1800
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Update"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin vkUserContolsXP.vkLabel vkLabel16 
               Height          =   255
               Left            =   240
               TabIndex        =   71
               Top             =   2760
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Address"
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
            Begin vkUserContolsXP.vkLabel vkLabel17 
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   3840
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "City"
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
            Begin vkUserContolsXP.vkLabel vkLabel18 
               Height          =   255
               Left            =   240
               TabIndex        =   69
               Top             =   4200
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Zipcode"
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
            Begin vkUserContolsXP.vkLabel vkLabel19 
               Height          =   255
               Left            =   240
               TabIndex        =   68
               Top             =   4560
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "State"
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
            Begin vkUserContolsXP.vkTextBox txtAddressAdmin 
               Height          =   975
               Left            =   2760
               TabIndex        =   67
               Top             =   2760
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1720
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
               Enabled         =   0   'False
               MultiLine       =   -1  'True
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtCityAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   66
               Top             =   3840
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtZipcodeAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   65
               Top             =   4200
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtStateAdmin 
               Height          =   255
               Left            =   2760
               TabIndex        =   64
               Top             =   4560
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnUp2Admin 
               Height          =   375
               Left            =   5400
               TabIndex        =   63
               Top             =   5040
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Update"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin vkUserContolsXP.vkLabel vkLabel20 
               Height          =   255
               Left            =   7080
               TabIndex        =   62
               Top             =   600
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Tel No (Home)"
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
            Begin vkUserContolsXP.vkLabel vkLabel21 
               Height          =   255
               Left            =   7080
               TabIndex        =   61
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Tel No (Mobile)"
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
            Begin vkUserContolsXP.vkLabel vkLabel22 
               Height          =   255
               Left            =   7080
               TabIndex        =   60
               Top             =   1320
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Position"
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
            Begin vkUserContolsXP.vkLabel vkLabel23 
               Height          =   255
               Left            =   7080
               TabIndex        =   59
               Top             =   1680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Commission"
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
            Begin vkUserContolsXP.vkTextBox txtTelHomeAdmin 
               Height          =   255
               Left            =   9360
               TabIndex        =   58
               Top             =   600
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtTelMobileAdmin 
               Height          =   255
               Left            =   9360
               TabIndex        =   57
               Top             =   960
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtPositionAdmin 
               Height          =   255
               Left            =   9360
               TabIndex        =   56
               Top             =   1320
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtCommissionAdmin 
               Height          =   255
               Left            =   9360
               TabIndex        =   55
               Top             =   1680
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnUp3Admin 
               Height          =   375
               Left            =   11040
               TabIndex        =   54
               Top             =   3360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Update"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
            Begin vkUserContolsXP.vkCommand btnDelStaff 
               Height          =   375
               Left            =   9960
               TabIndex        =   53
               Top             =   5040
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
               BackColor1      =   255
               BackColor2      =   8421631
               BackColorPushed1=   12632256
               Caption         =   "Delete"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16777215
               CustomStyle     =   0
            End
            Begin vkUserContolsXP.vkLabel vkLabel39 
               Height          =   255
               Left            =   7080
               TabIndex        =   52
               Top             =   2040
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Salary Scale"
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
            Begin vkUserContolsXP.vkLabel vkLabel40 
               Height          =   255
               Left            =   7080
               TabIndex        =   51
               Top             =   2760
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Hourly Rate"
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
            Begin vkUserContolsXP.vkTextBox txtSalaryScale2 
               Height          =   255
               Left            =   9360
               TabIndex        =   50
               Top             =   2040
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtHourlyRate2 
               Height          =   255
               Left            =   9360
               TabIndex        =   49
               Top             =   2760
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnEdit1 
               Height          =   375
               Left            =   3960
               TabIndex        =   48
               Top             =   1800
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Edit"
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
            Begin vkUserContolsXP.vkCommand btnEdit2 
               Height          =   375
               Left            =   3960
               TabIndex        =   47
               Top             =   5040
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Edit"
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
            Begin vkUserContolsXP.vkCommand btnEdit3 
               Height          =   375
               Left            =   9600
               TabIndex        =   46
               Top             =   3360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               Caption         =   "Edit"
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
         Begin vkUserContolsXP.vkFrame vkFrame2 
            Height          =   6015
            Left            =   -75000
            TabIndex        =   79
            Top             =   300
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   10610
            Caption         =   "Add New Staff"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox txtStaffIDInsert 
               Height          =   285
               Left            =   5520
               TabIndex        =   113
               Text            =   "Text1"
               Top             =   1680
               Visible         =   0   'False
               Width           =   855
            End
            Begin vkUserContolsXP.vkFrame frmUserDetailsInsert 
               Height          =   3855
               Left            =   120
               TabIndex        =   91
               Top             =   2040
               Visible         =   0   'False
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   6800
               BackColor1      =   14737632
               BackColor2      =   16777215
               Caption         =   "Details"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TitleColor1     =   14737632
               TitleColor2     =   4210752
               TitleGradient   =   2
               BorderColor     =   14737632
               Begin VB.ComboBox cmbStatusI 
                  Height          =   315
                  ItemData        =   "frmMain2.frx":4E22
                  Left            =   9240
                  List            =   "frmMain2.frx":4E2C
                  TabIndex        =   117
                  Text            =   "STATUS"
                  Top             =   2880
                  Width           =   3015
               End
               Begin vkUserContolsXP.vkLabel vkLabel42 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   116
                  Top             =   2880
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "State"
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
               Begin vkUserContolsXP.vkCommand btnInsert3 
                  Height          =   375
                  Left            =   9840
                  TabIndex        =   112
                  Top             =   3360
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   661
                  Caption         =   "Insert"
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
               Begin vkUserContolsXP.vkTextBox txtCommissionAdminI 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   111
                  Top             =   1800
                  Width           =   3015
                  _ExtentX        =   5318
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
               Begin vkUserContolsXP.vkTextBox txtPositionAdminI 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   110
                  Top             =   1440
                  Width           =   3015
                  _ExtentX        =   5318
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
               Begin vkUserContolsXP.vkTextBox txtTel_MobileAdminI 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   109
                  Top             =   1080
                  Width           =   3015
                  _ExtentX        =   5318
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
               Begin vkUserContolsXP.vkTextBox txtTel_HomeAdminI 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   108
                  Top             =   720
                  Width           =   3015
                  _ExtentX        =   5318
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
               Begin vkUserContolsXP.vkLabel vkLabel25 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   107
                  Top             =   1800
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Commission"
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
               Begin vkUserContolsXP.vkLabel vkLabel26 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   106
                  Top             =   1440
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Position"
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
               Begin vkUserContolsXP.vkLabel vkLabel27 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Tel No (Mobile)"
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
               Begin vkUserContolsXP.vkLabel vkLabel28 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   104
                  Top             =   720
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Tel No (Home)"
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
               Begin vkUserContolsXP.vkTextBox txtStateAdminI 
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   103
                  Top             =   2520
                  Width           =   3975
                  _ExtentX        =   7011
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
               Begin vkUserContolsXP.vkTextBox txtZipcodeAdminI 
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   102
                  Top             =   2160
                  Width           =   3975
                  _ExtentX        =   7011
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
               Begin vkUserContolsXP.vkTextBox txtCityAdminI 
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   101
                  Top             =   1800
                  Width           =   3975
                  _ExtentX        =   7011
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
               Begin vkUserContolsXP.vkTextBox txtAddressAdminI 
                  Height          =   975
                  Left            =   2640
                  TabIndex        =   100
                  Top             =   720
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   1720
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
                  MultiLine       =   -1  'True
                  LegendForeColor =   15695701
               End
               Begin vkUserContolsXP.vkLabel vkLabel29 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   99
                  Top             =   2520
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "State"
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
               Begin vkUserContolsXP.vkLabel vkLabel30 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   98
                  Top             =   2160
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Zipcode"
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
               Begin vkUserContolsXP.vkLabel vkLabel31 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   97
                  Top             =   1800
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "City"
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
               Begin vkUserContolsXP.vkLabel vkLabel32 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   96
                  Top             =   720
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Address"
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
               Begin vkUserContolsXP.vkLabel vkLabel37 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   95
                  Top             =   2160
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Salary Scale"
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
               Begin vkUserContolsXP.vkLabel vkLabel38 
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   94
                  Top             =   2520
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  BackColor       =   16777215
                  BackStyle       =   0
                  Caption         =   "Hourly Rate"
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
               Begin vkUserContolsXP.vkTextBox txtSalaryScale 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   93
                  Top             =   2160
                  Width           =   3015
                  _ExtentX        =   5318
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
               Begin vkUserContolsXP.vkTextBox txtHourlyRate 
                  Height          =   255
                  Left            =   9240
                  TabIndex        =   92
                  Top             =   2520
                  Width           =   3015
                  _ExtentX        =   5318
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
            End
            Begin VB.ComboBox cmbRole 
               Height          =   315
               ItemData        =   "frmMain2.frx":4E46
               Left            =   9960
               List            =   "frmMain2.frx":4E50
               TabIndex        =   90
               Text            =   "ROLE"
               Top             =   960
               Width           =   2415
            End
            Begin vkUserContolsXP.vkLabel vkLabel41 
               Height          =   255
               Left            =   7440
               TabIndex        =   89
               Top             =   960
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkTextBox txtPassword 
               Height          =   255
               Left            =   2760
               TabIndex        =   88
               Top             =   1320
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
            Begin vkUserContolsXP.vkLabel vkLabel36 
               Height          =   255
               Left            =   240
               TabIndex        =   87
               Top             =   1320
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel35 
               Height          =   255
               Left            =   240
               TabIndex        =   86
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel34 
               Height          =   255
               Left            =   7440
               TabIndex        =   85
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel33 
               Height          =   255
               Left            =   240
               TabIndex        =   84
               Top             =   960
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkTextBox txtNameAdminI 
               Height          =   255
               Left            =   2760
               TabIndex        =   83
               Top             =   600
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
            Begin vkUserContolsXP.vkTextBox txtUsernameAdminI 
               Height          =   255
               Left            =   9960
               TabIndex        =   82
               Top             =   600
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
            Begin vkUserContolsXP.vkTextBox txtEmailAdminI 
               Height          =   255
               Left            =   2760
               TabIndex        =   81
               Top             =   960
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
            Begin vkUserContolsXP.vkCommand btnInsert1 
               Height          =   375
               Left            =   9840
               TabIndex        =   80
               Top             =   1440
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               Caption         =   "Insert"
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4575
            Left            =   -75000
            TabIndex        =   114
            Top             =   1380
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   8070
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   6
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
         Begin vkUserContolsXP.vkCommand btnRetrieve 
            Height          =   375
            Left            =   -71400
            TabIndex        =   115
            Top             =   780
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Show Report"
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
         Begin vkUserContolsXP.vkCommand btnExportCSVStaff 
            Height          =   375
            Left            =   -68880
            TabIndex        =   204
            Top             =   780
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Export to Excel"
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
      Begin TabDlg.SSTab TabMyProfile 
         Height          =   6735
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BackColor       =   16766935
         TabCaption(0)   =   "Edit Profile"
         TabPicture(0)   =   "frmMain2.frx":4E62
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vkFrame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin vkUserContolsXP.vkFrame vkFrame1 
            Height          =   6255
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   11033
            Caption         =   "Edit Profile"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin vkUserContolsXP.vkLabel vkLabel2 
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel3 
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   960
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkLabel vkLabel4 
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   1320
               Width           =   2415
               _ExtentX        =   4260
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
            Begin vkUserContolsXP.vkTextBox txtName 
               Height          =   255
               Left            =   2760
               TabIndex        =   37
               Top             =   600
               Width           =   3975
               _ExtentX        =   7011
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
            Begin vkUserContolsXP.vkTextBox txtUsername 
               Height          =   255
               Left            =   2760
               TabIndex        =   36
               Top             =   960
               Width           =   3975
               _ExtentX        =   7011
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtEmail 
               Height          =   255
               Left            =   2760
               TabIndex        =   35
               Top             =   1320
               Width           =   3975
               _ExtentX        =   7011
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
            Begin vkUserContolsXP.vkCommand btnUpdateP1 
               Height          =   375
               Left            =   4200
               TabIndex        =   34
               Top             =   1800
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               Caption         =   "Update"
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
            Begin vkUserContolsXP.vkLabel vkLabel5 
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   2760
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Address"
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
            Begin vkUserContolsXP.vkLabel vkLabel6 
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   3840
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "City"
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
            Begin vkUserContolsXP.vkLabel vkLabel7 
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   4200
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Zipcode"
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
            Begin vkUserContolsXP.vkLabel vkLabel8 
               Height          =   255
               Left            =   240
               TabIndex        =   30
               Top             =   4560
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "State"
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
            Begin vkUserContolsXP.vkTextBox txtAddress 
               Height          =   975
               Left            =   2760
               TabIndex        =   29
               Top             =   2760
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1720
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
               MultiLine       =   -1  'True
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkTextBox txtCity 
               Height          =   255
               Left            =   2760
               TabIndex        =   28
               Top             =   3840
               Width           =   3975
               _ExtentX        =   7011
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
            Begin vkUserContolsXP.vkTextBox txtZipcode 
               Height          =   255
               Left            =   2760
               TabIndex        =   27
               Top             =   4200
               Width           =   3975
               _ExtentX        =   7011
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
            Begin vkUserContolsXP.vkTextBox txtState 
               Height          =   255
               Left            =   2760
               TabIndex        =   26
               Top             =   4560
               Width           =   3975
               _ExtentX        =   7011
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
            Begin vkUserContolsXP.vkCommand btnUpdateP2 
               Height          =   375
               Left            =   4200
               TabIndex        =   25
               Top             =   5040
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               Caption         =   "Update"
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
            Begin vkUserContolsXP.vkLabel vkLabel9 
               Height          =   255
               Left            =   7080
               TabIndex        =   24
               Top             =   600
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Tel No (Home)"
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
            Begin vkUserContolsXP.vkLabel vkLabel10 
               Height          =   255
               Left            =   7080
               TabIndex        =   23
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Tel No (Mobile)"
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
            Begin vkUserContolsXP.vkLabel vkLabel11 
               Height          =   255
               Left            =   7080
               TabIndex        =   22
               Top             =   1320
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Position"
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
            Begin vkUserContolsXP.vkLabel vkLabel12 
               Height          =   255
               Left            =   7080
               TabIndex        =   21
               Top             =   1680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Commission"
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
            Begin vkUserContolsXP.vkTextBox txtTel_Home 
               Height          =   255
               Left            =   9360
               TabIndex        =   20
               Top             =   600
               Width           =   3015
               _ExtentX        =   5318
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
            Begin vkUserContolsXP.vkTextBox txtTel_Mobile 
               Height          =   255
               Left            =   9360
               TabIndex        =   19
               Top             =   960
               Width           =   3015
               _ExtentX        =   5318
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
            Begin vkUserContolsXP.vkTextBox txtPosition 
               Height          =   255
               Left            =   9360
               TabIndex        =   18
               Top             =   1320
               Width           =   3015
               _ExtentX        =   5318
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
            Begin vkUserContolsXP.vkTextBox txtCommission 
               Height          =   255
               Left            =   9360
               TabIndex        =   17
               Top             =   1680
               Width           =   3015
               _ExtentX        =   5318
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
               Enabled         =   0   'False
               LegendForeColor =   15695701
            End
            Begin vkUserContolsXP.vkCommand btnUpdateP3 
               Height          =   375
               Left            =   9960
               TabIndex        =   16
               Top             =   2160
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
               Caption         =   "Update"
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
      Begin TabDlg.SSTab tabInventoryAdmin 
         Height          =   6735
         Left            =   120
         TabIndex        =   160
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Add Inventory"
         TabPicture(0)   =   "frmMain2.frx":4E7E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "btnNextInvAdmin"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnPrevInvAdmin"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vkFrame5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "btnDelInvAdmin"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtInvIDAdmin"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "btnUpdateInvAdmin"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "View Report"
         TabPicture(1)   =   "frmMain2.frx":4E9A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cdInventory"
         Tab(1).Control(1)=   "cmbMthInv"
         Tab(1).Control(2)=   "cmbYrInv"
         Tab(1).Control(3)=   "btnYearInv"
         Tab(1).Control(4)=   "btnMonthInv"
         Tab(1).Control(5)=   "DataGrid6"
         Tab(1).Control(6)=   "btnShowReportInv"
         Tab(1).Control(7)=   "btnExportToXLSInv"
         Tab(1).ControlCount=   8
         Begin MSComDlg.CommonDialog cdInventory 
            Left            =   -69960
            Top             =   6120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox cmbMthInv 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4EB6
            Left            =   -74880
            List            =   "frmMain2.frx":4EDE
            TabIndex        =   206
            Text            =   "Month"
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox cmbYrInv 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain2.frx":4F12
            Left            =   -73200
            List            =   "frmMain2.frx":4F58
            TabIndex        =   205
            Text            =   "Year"
            Top             =   480
            Width           =   1695
         End
         Begin vkUserContolsXP.vkCommand btnUpdateInvAdmin 
            Height          =   495
            Left            =   240
            TabIndex        =   198
            Top             =   4200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            Caption         =   "Update Inventory"
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
         Begin VB.TextBox txtInvIDAdmin 
            Height          =   285
            Left            =   3120
            TabIndex        =   162
            Top             =   5520
            Visible         =   0   'False
            Width           =   2055
         End
         Begin vkUserContolsXP.vkCommand btnDelInvAdmin 
            Height          =   495
            Left            =   3120
            TabIndex        =   161
            Top             =   4800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BackColor1      =   8421631
            BackColorPushed1=   12632319
            BackGradient    =   0
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   0
         End
         Begin MSDataGridLib.DataGrid DataGrid4 
            Height          =   5655
            Left            =   5400
            TabIndex        =   163
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   9975
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
         Begin vkUserContolsXP.vkFrame vkFrame5 
            Height          =   3735
            Left            =   240
            TabIndex        =   164
            Top             =   240
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
            Begin vkUserContolsXP.vkLabel vkLabel58 
               Height          =   255
               Left            =   240
               TabIndex        =   175
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
            Begin vkUserContolsXP.vkLabel vkLabel57 
               Height          =   255
               Left            =   240
               TabIndex        =   174
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
            Begin vkUserContolsXP.vkLabel vkLabel56 
               Height          =   255
               Left            =   240
               TabIndex        =   173
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
            Begin vkUserContolsXP.vkLabel vkLabel55 
               Height          =   255
               Left            =   240
               TabIndex        =   172
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
            Begin vkUserContolsXP.vkLabel vkLabel54 
               Height          =   255
               Left            =   240
               TabIndex        =   171
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
            Begin vkUserContolsXP.vkTextBox txtItemDescAdmin 
               Height          =   255
               Left            =   2040
               TabIndex        =   170
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
            Begin vkUserContolsXP.vkTextBox txtPriceAdmin 
               Height          =   255
               Left            =   2040
               TabIndex        =   169
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
            Begin vkUserContolsXP.vkTextBox txtQISAdmin 
               Height          =   255
               Left            =   2040
               TabIndex        =   168
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
            Begin vkUserContolsXP.vkTextBox txtMinOrderAdmin 
               Height          =   255
               Left            =   2040
               TabIndex        =   167
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
            Begin vkUserContolsXP.vkTextBox txtTypeOfItemAdmin 
               Height          =   255
               Left            =   2040
               TabIndex        =   166
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
            Begin vkUserContolsXP.vkCommand btnAddInvAdmin 
               Height          =   375
               Left            =   2040
               TabIndex        =   165
               Top             =   3120
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               Caption         =   "ADD INVENTORY"
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
         End
         Begin vkUserContolsXP.vkCommand btnPrevInvAdmin 
            Height          =   495
            Left            =   3120
            TabIndex        =   176
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Previous"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin vkUserContolsXP.vkCommand btnNextInvAdmin 
            Height          =   495
            Left            =   4200
            TabIndex        =   177
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Next"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin vkUserContolsXP.vkCommand btnYearInv 
            Height          =   375
            Left            =   -73200
            TabIndex        =   207
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Yearly"
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
         Begin vkUserContolsXP.vkCommand btnMonthInv 
            Height          =   375
            Left            =   -74880
            TabIndex        =   208
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Monthly"
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
         Begin MSDataGridLib.DataGrid DataGrid6 
            Height          =   4575
            Left            =   -75000
            TabIndex        =   209
            Top             =   1080
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   8070
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
            RowDividerStyle =   6
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
         Begin vkUserContolsXP.vkCommand btnShowReportInv 
            Height          =   375
            Left            =   -71400
            TabIndex        =   210
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Show Report"
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
         Begin vkUserContolsXP.vkCommand btnExportToXLSInv 
            Height          =   375
            Left            =   -68880
            TabIndex        =   211
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            Caption         =   "Export to Excel"
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
      Begin TabDlg.SSTab tabInventoryStaff 
         Height          =   6735
         Left            =   120
         TabIndex        =   142
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   11880
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Add Inventory"
         TabPicture(0)   =   "frmMain2.frx":4FE0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "btnNextInvStaff"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "btnPrevInvStaff"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vkFrame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "DataGrid3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtInvID"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "btnDelInvStaff"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "btnUpdateInvStaff"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         Begin vkUserContolsXP.vkCommand btnUpdateInvStaff 
            Height          =   495
            Left            =   240
            TabIndex        =   199
            Top             =   4200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            Caption         =   "Update Inventory"
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
         Begin vkUserContolsXP.vkCommand btnDelInvStaff 
            Height          =   495
            Left            =   3120
            TabIndex        =   159
            Top             =   4800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            BackColor1      =   8421631
            BackColorPushed1=   12632319
            BackGradient    =   0
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   0
         End
         Begin VB.TextBox txtInvID 
            Height          =   285
            Left            =   3120
            TabIndex        =   158
            Top             =   5520
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   5655
            Left            =   5400
            TabIndex        =   155
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   9975
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
         Begin vkUserContolsXP.vkFrame vkFrame4 
            Height          =   3735
            Left            =   240
            TabIndex        =   143
            Top             =   240
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
            Begin vkUserContolsXP.vkCommand btnAddInvStaff 
               Height          =   375
               Left            =   2040
               TabIndex        =   154
               Top             =   3120
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   661
               Caption         =   "ADD INVENTORY"
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
            Begin vkUserContolsXP.vkTextBox txtTypeOfITemStaff 
               Height          =   255
               Left            =   2040
               TabIndex        =   153
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
            Begin vkUserContolsXP.vkTextBox txtMinOrderStaff 
               Height          =   255
               Left            =   2040
               TabIndex        =   152
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
            Begin vkUserContolsXP.vkTextBox txtQISStaff 
               Height          =   255
               Left            =   2040
               TabIndex        =   151
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
            Begin vkUserContolsXP.vkTextBox txtPriceStaff 
               Height          =   255
               Left            =   2040
               TabIndex        =   150
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
            Begin vkUserContolsXP.vkTextBox txtItemDescStaff 
               Height          =   255
               Left            =   2040
               TabIndex        =   149
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
            Begin vkUserContolsXP.vkLabel vkLabel53 
               Height          =   255
               Left            =   240
               TabIndex        =   148
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
            Begin vkUserContolsXP.vkLabel vkLabel52 
               Height          =   255
               Left            =   240
               TabIndex        =   147
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
            Begin vkUserContolsXP.vkLabel vkLabel51 
               Height          =   255
               Left            =   240
               TabIndex        =   146
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
            Begin vkUserContolsXP.vkLabel vkLabel50 
               Height          =   255
               Left            =   240
               TabIndex        =   145
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
            Begin vkUserContolsXP.vkLabel vkLabel49 
               Height          =   255
               Left            =   240
               TabIndex        =   144
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
         Begin vkUserContolsXP.vkCommand btnPrevInvStaff 
            Height          =   495
            Left            =   3120
            TabIndex        =   156
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Previous"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
         Begin vkUserContolsXP.vkCommand btnNextInvStaff 
            Height          =   495
            Left            =   4200
            TabIndex        =   157
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BackColor1      =   16761024
            BackColorPushed1=   16761024
            BackGradient    =   0
            Caption         =   "Next"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   4210752
            BorderColor     =   8421504
            DrawFocus       =   0   'False
            DrawMouseInRect =   0   'False
            CustomStyle     =   3
         End
      End
   End
   Begin vkUserContolsXP.vkFrame frmSales 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1931
      BackColor1      =   14737632
      BackColor2      =   12632256
      Caption         =   "Menu - Staff Directory"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      TitleColor1     =   16761024
      TitleColor2     =   16761024
      BorderColor     =   16761024
      Begin vkUserContolsXP.vkCommand btnPunchOut2 
         Height          =   615
         Index           =   0
         Left            =   11160
         TabIndex        =   11
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BackColor1      =   8421631
         BackColor2      =   255
         BackColorPushed1=   255
         BackGradient    =   1
         Caption         =   "Punch Out"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   16777215
         Picture         =   "frmMain2.frx":4FFC
         CustomStyle     =   0
         MouseHoverPicture=   "frmMain2.frx":6950
      End
      Begin vkUserContolsXP.vkCommand btnPunchIn 
         Height          =   615
         Left            =   9600
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16761024
         BackColor2      =   16711680
         BackColorPushed1=   16711680
         Caption         =   "Punch In"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   16777215
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand btnInventoryStaff 
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Inventory"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkCommand btnSalesStaff 
         Height          =   615
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Sales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkCommand btnMyProfile 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "My Profile"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
   End
   Begin vkUserContolsXP.vkFrame frmAdmin 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   1931
      BackColor1      =   14737632
      BackColor2      =   12632256
      Caption         =   "Menu - Human Resource"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      TitleColor1     =   16761024
      TitleColor2     =   16761024
      BorderColor     =   16761024
      Begin vkUserContolsXP.vkCommand btnPunchOut 
         Height          =   615
         Index           =   1
         Left            =   11160
         TabIndex        =   13
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         BackColor1      =   8421631
         BackColor2      =   255
         BackColorPushed1=   255
         BackGradient    =   1
         Caption         =   "Punch Out"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   16777215
         Picture         =   "frmMain2.frx":82A4
         CustomStyle     =   0
         MouseHoverPicture=   "frmMain2.frx":9BF8
      End
      Begin vkUserContolsXP.vkCommand btnPunchInAdmin 
         Height          =   615
         Left            =   9600
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16761024
         BackColor2      =   16711680
         BackColorPushed1=   16711680
         Caption         =   "Punch In"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   16777215
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkCommand btnStaff 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Staff"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkCommand btnInventory 
         Height          =   615
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Inventory"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
      Begin vkUserContolsXP.vkCommand btnSales 
         Height          =   615
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Sales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
      End
   End
   Begin vkUserContolsXP.vkScrollContainer vkScrollContainer2 
      Height          =   8775
      Left            =   -120
      TabIndex        =   219
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   15478
   End
End
Attribute VB_Name = "frmMain2"
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
Dim FExists As Boolean

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

Private Sub btnAddInvAdmin_Click()
    Dim rs1
    
    OpenServer
    
    If IsNumeric(txtPriceAdmin.Text) = True And IsNumeric(txtQISAdmin.Text) And IsNumeric(txtMinOrderAdmin.Text) Then
    
        conn.Execute "INSERT INTO inventory(item_desc,price,quantity_instock,minimum_order_quantity,type_of_item,staffID)" _
        & " VALUES ( " _
        & " '" & txtItemDescAdmin.Text & "' " _
        & " ,'" & txtPriceAdmin.Text & "' " _
        & " ,'" & txtQISAdmin.Text & "' " _
        & " ,'" & txtMinOrderAdmin.Text & "' " _
        & " ,'" & txtTypeOfItemAdmin.Text & "' " _
        & " ,'" & staff_id & "')", , adExecuteNoRecords
        
        MsgBox ("Record saved")
        
        txtItemDescAdmin.Text = ""
        txtPriceAdmin.Text = ""
        txtQISAdmin.Text = ""
        txtMinOrderAdmin.Text = ""
        txtTypeOfItemAdmin.Text = ""
    
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.CursorType = adOpenStatic
        rs1.LockType = adLockReadOnly
        rs1.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  " _
        & "FROM inventory ", conn
        Set DataGrid4.DataSource = rs1
    Else
        MsgBox ("Please put numeric value only for Price or Quantity In Stock or Minimum Order Quantity")
    End If

End Sub

Private Sub btnAddInvStaff_Click()
    Dim rs1
    
    OpenServer
    
    If IsNumeric(txtPriceStaff.Text) = True And IsNumeric(txtQISStaff.Text) And IsNumeric(txtMinOrderStaff.Text) Then
    
        conn.Execute "INSERT INTO inventory(item_desc,price,quantity_instock,minimum_order_quantity,type_of_item,staffID)" _
        & " VALUES ( " _
        & " '" & txtItemDescStaff.Text & "' " _
        & " ,'" & txtPriceStaff.Text & "' " _
        & " ,'" & txtQISStaff.Text & "' " _
        & " ,'" & txtMinOrderStaff.Text & "' " _
        & " ,'" & txtTypeOfITemStaff.Text & "' " _
        & " ,'" & staff_id & "')", , adExecuteNoRecords
        
        MsgBox ("Record saved")
        
        txtItemDescStaff.Text = ""
        txtPriceStaff.Text = ""
        txtQISStaff.Text = ""
        txtMinOrderStaff.Text = ""
        txtTypeOfITemStaff.Text = ""
    
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.CursorType = adOpenStatic
        rs1.LockType = adLockReadOnly
        rs1.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  " _
        & "FROM inventory ", conn
        Set DataGrid3.DataSource = rs1
    Else
        MsgBox ("Please put numeric value only for Price or Quantity In Stock or Minimum Order Quantity")
    End If
End Sub

Private Sub btnAddItemAdmin_Click()
    Dim QuantityLeft As Double
    Dim price As Double
    Dim SalesID As Long
    Dim totalPrice As Double
    Dim MinOrderQuantity As Double
    Dim ItemDesc As String
    Dim TotalingPrice As Double
    Dim rs1
    Dim iID As Long
    Dim aSession As String
    Dim aSessionDay As String
    
    OpenServer
    
    If IsNumeric(txtQuantityAdmin.Text) = True And CInt(txtQuantityAdmin.Text) <> 0 Then
    
        Set rsQuery1 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM inventory where id='" & txtItemIDAdmin.Text & "' "
        Set rsQuery1 = conn.Execute(SQLText)
        
        If Not rsQuery1.EOF Then
            QuantityLeft = rsQuery1("quantity_instock")
            MinOrderQuantity = rsQuery1("minimum_order_quantity")
            price = rsQuery1("price")
            ItemDesc = rsQuery1("item_desc")
            iID = rsQuery1("id")
            
            txtItemDescAddAdmin.Text = ItemDesc
            TotalingPrice = price * CLng(txtQuantityAdmin.Text)
            
            If CLng(txtQuantityAdmin.Text) <= QuantityLeft Then
                If QuantityLeft <> MinOrderQuantity Then
                    conn.Execute "INSERT INTO item_sale(quantity_sold,salesID,total_price,item_desc,session,price,staffID)" _
                    & " VALUES ( " _
                    & " '" & txtQuantityAdmin.Text & "' " _
                    & " ,'" & txtItemIDAdmin.Text & "' " _
                    & " ,'" & TotalingPrice & "' " _
                    & " ,'" & ItemDesc & "' " _
                    & " ,'" & txtSessionAdmin.Text & "' " _
                    & " ,'" & price & "' " _
                    & " ,'" & staff_id & "')", , adExecuteNoRecords
                    
                    conn.Execute "UPDATE inventory set quantity_instock='" & QuantityLeft - CLng(txtQuantityAdmin.Text) & "'" _
                    & " WHERE id='" & iID & "'", , adExecuteNoRecords
                    
                    Set rs1 = New ADODB.Recordset
                    rs1.CursorLocation = adUseClient
                    rs1.CursorType = adOpenStatic
                    rs1.LockType = adLockReadOnly
                    rs1.Open "SELECT item_desc as Item, quantity_sold as Quantity, price as Price, total_price as Total_Price " _
                    & "FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
                    Set DataGrid5.DataSource = rs1
                    
                    MsgBox ("New Item add")
                 ElseIf QuantityLeft <= MinOrderQuantity Then
                    conn.Execute "INSERT INTO item_sale(quantity_sold,salesID,total_price,item_desc,session,price,staffID)" _
                    & " VALUES ( " _
                    & " '" & txtQuantityAdmin.Text & "' " _
                    & " ,'" & txtItemIDAdmin.Text & "' " _
                    & " ,'" & TotalingPrice & "' " _
                    & " ,'" & ItemDesc & "' " _
                    & " ,'" & txtSessionAdmin.Text & "' " _
                    & " ,'" & price & "' " _
                    & " ,'" & staff_id & "')", , adExecuteNoRecords
                    
                    conn.Execute "UPDATE inventory set quantity_instock='" & QuantityLeft - CLng(txtQuantityAdmin.Text) & "'" _
                    & " WHERE id='" & iID & "'", , adExecuteNoRecords
                    
                    Set rs1 = New ADODB.Recordset
                    rs1.CursorLocation = adUseClient
                    rs1.CursorType = adOpenStatic
                    rs1.LockType = adLockReadOnly
                    rs1.Open "SELECT item_desc as Item, quantity_sold as Quantity,price as Price, total_price as Total_Price " _
                    & "FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
                    Set DataGrid5.DataSource = rs1
                    
                    MsgBox ("New Item add")
                    MsgBox ("Quantity request is reach to minimum Quantity in stock. Please re-stock")
                 End If
            Else
                MsgBox ("Quantity request is larger than Quantity in stock.")
                txtQuantityAdmin.Text = ""
                txtItemIDAdmin.Text = ""
                txtItemDescAdmin.Text = ""
            End If
        Else
            MsgBox ("Invalid Item ID!!!")
            txtQuantityAdmin.Text = ""
            txtItemIDAdmin.Text = ""
            txtItemDescAdmin.Text = ""
        End If
        
        Set rsQuery2 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM item_sale where session='" & txtSessionAdmin.Text & "' "
        Set rsQuery2 = conn.Execute(SQLText)
        
        totalPrice = 0
        
        Do While Not rsQuery2.EOF
            price = rsQuery2("total_price")
            totalPrice = totalPrice + price
        rsQuery2.MoveNext
        Loop
        
        lblTotalPriceAdmin.Caption = "RM " & totalPrice
    Else
        MsgBox ("Please enter only numeric for Quantity or Quantity must be greater than 0")
    End If
End Sub

Private Sub btnAddItemStaff_Click()
    Dim QuantityLeft As Double
    Dim price As Double
    Dim SalesID As Long
    Dim totalPrice As Double
    Dim MinOrderQuantity As Double
    Dim ItemDesc As String
    Dim TotalingPrice As Double
    Dim rs1
    Dim iID As Long
    
    OpenServer
    
    If IsNumeric(txtQuantityAddS.Text) = True And CInt(txtQuantityAddS.Text) <> 0 Then
        Set rsQuery1 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM inventory where id='" & txtItemIDAddS.Text & "' "
        Set rsQuery1 = conn.Execute(SQLText)
        
        If Not rsQuery1.EOF Then
            QuantityLeft = rsQuery1("quantity_instock")
            MinOrderQuantity = rsQuery1("minimum_order_quantity")
            price = rsQuery1("price")
            ItemDesc = rsQuery1("item_desc")
            iID = rsQuery1("id")
            
            txtItemDescAddS.Text = ItemDesc
            TotalingPrice = price * CLng(txtQuantityAddS.Text)
            
            If CLng(txtQuantityAddS.Text) <= QuantityLeft Then
                If QuantityLeft <> MinOrderQuantity Then
                    conn.Execute "INSERT INTO item_sale(quantity_sold,salesID,total_price,item_desc,session,price,staffID)" _
                    & " VALUES ( " _
                    & " '" & txtQuantityAddS.Text & "' " _
                    & " ,'" & txtItemIDAddS.Text & "' " _
                    & " ,'" & TotalingPrice & "' " _
                    & " ,'" & ItemDesc & "' " _
                    & " ,'" & txtSessionStaff.Text & "' " _
                    & " ,'" & price & "' " _
                    & " ,'" & staff_id & "')", , adExecuteNoRecords
                    
                    conn.Execute "UPDATE inventory set quantity_instock='" & QuantityLeft - CLng(txtQuantityAddS.Text) & "'" _
                    & " WHERE id='" & iID & "'", , adExecuteNoRecords
                    
                    Set rs1 = New ADODB.Recordset
                    rs1.CursorLocation = adUseClient
                    rs1.CursorType = adOpenStatic
                    rs1.LockType = adLockReadOnly
                    rs1.Open "SELECT item_desc as Item, quantity_sold as Quantity,price as Price, total_price as Total_Price " _
                    & "FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
                    Set DataGrid2.DataSource = rs1
                    
                    MsgBox ("New Item add")
                 ElseIf QuantityLeft <= MinOrderQuantity Then
                    conn.Execute "INSERT INTO item_sale(quantity_sold,salesID,total_price,item_desc,session,price,staffID)" _
                    & " VALUES ( " _
                    & " '" & txtQuantityAddS.Text & "' " _
                    & " ,'" & txtItemIDAddS.Text & "' " _
                    & " ,'" & TotalingPrice & "' " _
                    & " ,'" & ItemDesc & "' " _
                    & " ,'" & txtSessionStaff.Text & "' " _
                    & " ,'" & price & "' " _
                    & " ,'" & staff_id & "')", , adExecuteNoRecords
                    
                    conn.Execute "UPDATE inventory set quantity_instock='" & QuantityLeft - CLng(txtQuantityAddS.Text) & "'" _
                    & " WHERE id='" & iID & "'", , adExecuteNoRecords
                    
                    Set rs1 = New ADODB.Recordset
                    rs1.CursorLocation = adUseClient
                    rs1.CursorType = adOpenStatic
                    rs1.LockType = adLockReadOnly
                    rs1.Open "SELECT item_desc as Item, quantity_sold as Quantity, price as Price, total_price as Total_Price " _
                    & "FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
                    Set DataGrid2.DataSource = rs1
                    
                    MsgBox ("New Item add")
                    MsgBox ("Quantity request is reach to minimum Quantity in stock. Please re-Stock")
                 End If
            Else
                MsgBox ("Quantity request is larger than Quantity in stock.")
                txtQuantityAddS.Text = ""
                txtItemIDAddS.Text = ""
                txtItemDescAddS.Text = ""
            End If
        Else
            MsgBox ("Invalid Item ID!!!")
            txtQuantityAddS.Text = ""
            txtItemIDAddS.Text = ""
            txtItemDescAddS.Text = ""
        End If
        
        Set rsQuery2 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM item_sale where session='" & txtSessionStaff.Text & "' "
        Set rsQuery2 = conn.Execute(SQLText)
        
        totalPrice = 0
        
        Do While Not rsQuery2.EOF
            price = rsQuery2("total_price")
            totalPrice = totalPrice + price
        rsQuery2.MoveNext
        Loop
        
        lblTotPrice.Caption = "RM " & totalPrice

    Else
        MsgBox ("Please enter only numeric for Quantity or Quantity must be greater than o")
    End If
End Sub

Private Sub btnDeleteSalesStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    Dim quantity As Long
    Dim iID As Long
    Dim QuantitySold As Long
    Dim SalesID As Long
    
    OpenServer
    
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where id='" & txtItemIDStaff.Text & "' "
    Set rsQuery2 = conn.Execute(SQLText)
    
    If Not rsQuery2.EOF Then
        QuantitySold = rsQuery2("quantity_sold")
        SalesID = rsQuery2("salesID")
    End If
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM inventory where id='" & SalesID & "' "
    Set rsQuery1 = conn.Execute(SQLText)
    
    If Not rsQuery1.EOF Then
        quantity = rsQuery1("quantity_instock")
        iID = rsQuery1("id")
    End If
    
    conn.Execute "UPDATE inventory set quantity_instock='" & quantity + QuantitySold & "'" _
    & " WHERE id='" & iID & "'", , adExecuteNoRecords
    
    conn.Execute "DELETE FROM item_sale Where id = '" & txtItemIDStaff.Text & "'"
    
    Rx = Rx - 1
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT * FROM item_sale", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid2.DataSource = rs
    
    txtItemIDStaff.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where session='" & txtSessionStaff.Text & "' "
    Set rsQuery2 = conn.Execute(SQLText)
    
    Dim totalPrice As Double
    Dim price As Double
    
    totalPrice = 0
    
    Do While Not rsQuery2.EOF
        price = rsQuery2("total_price")
        totalPrice = totalPrice + price
        rsQuery2.MoveNext
    Loop
    
    lblTotPrice.Caption = "RM " & totalPrice
    
    Set DataGrid2.DataSource = rs2
        
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID, item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
    Set DataGrid2.DataSource = rs2
End Sub

Private Sub btnDelInvAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    conn.Execute "DELETE FROM inventory Where id = '" & txtInvIDAdmin.Text & "'"
    Rx = Rx - 1
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT * FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid4.DataSource = rs
    
    txtInvIDAdmin.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set DataGrid4.DataSource = rs2
        
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item FROM inventory", conn
    Set DataGrid4.DataSource = rs2
End Sub

Private Sub btnDelInvStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    conn.Execute "DELETE FROM inventory Where id = '" & txtInvID.Text & "'"
    Rx = Rx - 1
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT * FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid3.DataSource = rs
    
    txtInvID.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set DataGrid3.DataSource = rs2
        
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
    Set DataGrid3.DataSource = rs2
End Sub

Private Sub btnDelItemAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    Dim quantity As Long
    Dim iID As Long
    Dim QuantitySold As Long
    Dim SalesID As Long
    
    OpenServer
    
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where id='" & txtIIDAdmin.Text & "' "
    Set rsQuery2 = conn.Execute(SQLText)
    
    If Not rsQuery2.EOF Then
        QuantitySold = rsQuery2("quantity_sold")
        SalesID = rsQuery2("salesID")
    End If
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM inventory where id='" & SalesID & "' "
    Set rsQuery1 = conn.Execute(SQLText)
    
    If Not rsQuery1.EOF Then
        quantity = rsQuery1("quantity_instock")
        iID = rsQuery1("id")
    End If
    
    conn.Execute "UPDATE inventory set quantity_instock='" & quantity + QuantitySold & "'" _
    & " WHERE id='" & iID & "'", , adExecuteNoRecords
    
    conn.Execute "DELETE FROM item_sale Where id = '" & txtIIDAdmin.Text & "'"
    
    Rx = Rx - 1
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT * FROM item_sale", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid5.DataSource = rs
    
    txtIIDAdmin.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where session='" & txtSessionAdmin.Text & "' "
    Set rsQuery2 = conn.Execute(SQLText)
    
    Dim totalPrice As Double
    Dim price As Double
    
    totalPrice = 0
    
    Do While Not rsQuery2.EOF
        price = rsQuery2("total_price")
        totalPrice = totalPrice + price
    rsQuery2.MoveNext
    Loop
    
    lblTotalPriceAdmin.Caption = "RM " & totalPrice
    
    Set DataGrid5.DataSource = rs2
        
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID, item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
    Set DataGrid5.DataSource = rs2
End Sub

Private Sub btnDelStaff_Click()
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    Else
        conn.Execute "DELETE FROM staff WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
        conn.Execute "DELETE FROM staff_address WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
        conn.Execute "DELETE FROM staff_details WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
        frmUserProfile.Visible = False
        txtUsernameFind.Text = ""
    End If
End Sub

Private Sub btnEdit1_Click()
    txtNameAdmin.Enabled = True
    txtUsernameAdmin.Enabled = True
    txtEmailAdmin.Enabled = True
    btnUp1Admin.Enabled = True
End Sub

Private Sub btnEdit2_Click()
    txtAddressAdmin.Enabled = True
    txtCityAdmin.Enabled = True
    txtStateAdmin.Enabled = True
    txtZipcodeAdmin.Enabled = True
    btnUp2Admin.Enabled = True
End Sub

Private Sub btnEdit3_Click()
    cmbStatusU.Enabled = True
    txtTelHomeAdmin.Enabled = True
    txtTelMobileAdmin.Enabled = True
    txtPositionAdmin.Enabled = True
    txtCommissionAdmin.Enabled = True
    txtSalaryScale2.Enabled = True
    btnUp3Admin.Enabled = True
End Sub

Private Sub btnExportCSVStaff_Click()
    Dim fieldnum As Integer
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim daAnswer As String
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    Dim l As Long
    
    lenMonth = Len(cmbMonth.Text)
    lenYear = Len(cmbYear.Text)
    
    If cmbMonth.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYear.Text
    ElseIf cmbMonth.Enabled = True And cmbYear.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYear.Text & "-" & cmbMonth.Text
    End If

    With cdExportStaff
      .DialogTitle = "Export to CSV"
      .Filter = "Excel Import File (*.xls)|*.xls"
      .FileName = "Staff" & Format(Now, "YYYYMMDDHHNNSS")
      .ShowSave
    End With
    
    FileExists (cdExportStaff.FileName)
    If FExists = True Then
      MsgBox "File Exist!!!"
      btnExportCSVStaff_Click
    End If
    
    Set xlApp = Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT " _
    & "s.name as Name," _
    & "sd.status as Status," _
    & "((sa.hours_working * sd.salary_scale)/sd.hourly_rate) as Salary, " _
    & "sa.submit_date as Date " _
    & "FROM staff s,staff_details sd, staff_attendance sa WHERE left(sa.submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "' " _
    & " AND sd.staffID=s.staffID and sa.staffID=s.staffID"
    Set rsQuery1 = conn.Execute(SQLText)
            
    totSalary = 0
    l = 1
    xlSheet.Cells(1, 1) = "Name"
    xlSheet.Cells(1, 2) = "Salary"
    xlSheet.Cells(1, 3) = "Status"
    xlSheet.Cells(1, 4) = "Date"
    
    Do While Not rsQuery1.EOF
            l = l + 1
            Salary = rsQuery1("Salary")
            xlSheet.Cells(l, 1) = CStr("" & rsQuery1("Name") & "")
            xlSheet.Cells(l, 2) = CStr("RM" & rsQuery1("Salary") & "")
            xlSheet.Cells(l, 3) = CStr("" & rsQuery1("Status") & "")
            xlSheet.Cells(l, 4) = CStr("" & rsQuery1("Date") & "")
            totSalary = totSalary + Salary
        rsQuery1.MoveNext
    Loop
    
    xlSheet.Cells(l + 1, 1) = "Total Price"
    xlSheet.Cells(l + 1, 2) = ""
    xlSheet.Cells(l + 1, 3) = ""
    xlSheet.Cells(l + 1, 4) = "RM" & totSalary
    
    xlBook.SaveAs cdExportStaff.FileName
    xlBook.Close
    MsgBox ("Export Success")

End Sub

Function FileExists(ByVal FileName As String)

   Dim Exists As Integer
   
   On Local Error Resume Next 'If some problem continue, code handles problems inherintly
   Exists = Len(Dir(FileName$)) 'Dir returns either a null string (len 0) or a filename
   On Local Error GoTo 0
 If Exists = 0 Then 'Null string?
    FileExists = False
    FExists = False
Else
    FileExists = True
    FExists = True
End If
End Function

Private Sub btnExportToXLSInv_Click()
    Dim fieldnum As Integer
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim daAnswer As String
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    Dim l As Long
    
    lenMonth = Len(cmbMthInv.Text)
    lenYear = Len(cmbYrInv.Text)
    
    If cmbMthInv.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYrInv.Text
    ElseIf cmbMthInv.Enabled = True And cmbYrInv.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYrInv.Text & "-" & cmbMthInv.Text
    End If
    
    With cdInventory
      .DialogTitle = "Export to XLS"
      .Filter = "Excel Import File (*.xls)|*.xls"
      .FileName = "Inventory" & Format(Now, "YYYYMMDDHHNNSS")
      .ShowSave
    End With
    
    FileExists (cdInventory.FileName)
    If FExists = True Then
      MsgBox "File Exist!!!"
      btnExportToXLSInv_Click
    End If
    
    Set xlApp = Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item FROM inventory WHERE left(submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "' "
    Set rsQuery1 = conn.Execute(SQLText)
            
    totSalary = 0
    l = 1
    xlSheet.Cells(1, 1) = "Item"
    xlSheet.Cells(1, 2) = "Quantity"
    xlSheet.Cells(1, 3) = "Min Quantity"
    xlSheet.Cells(1, 4) = "Type Of Item"
    
    Do While Not rsQuery1.EOF
            l = l + 1
            xlSheet.Cells(l, 1) = CStr("" & rsQuery1("Item") & "")
            xlSheet.Cells(l, 2) = CStr("" & rsQuery1("Quantity") & "")
            xlSheet.Cells(l, 3) = CStr("" & rsQuery1("Min_Quantity") & "")
            xlSheet.Cells(l, 4) = CStr("" & rsQuery1("Type_Item") & "")
            
        rsQuery1.MoveNext
    Loop
    rsQuery1.Close
    
    xlBook.SaveAs cdInventory.FileName
    xlBook.Close
    MsgBox ("Export Success")
End Sub

Private Sub btnExportXLSSales_Click()
    Dim fieldnum As Integer
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim daAnswer As String
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    Dim l As Long
    Dim totalPrice As Double
    Dim PriceT As Double
    
    lenMonth = Len(cmbMthSales.Text)
    lenYear = Len(cmbYrSales.Text)
    
    If cmbMthSales.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYrSales.Text
    ElseIf cmbMthSales.Enabled = True And cmbYrSales.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYrSales.Text & "-" & cmbMthSales.Text
    End If
  
    cdExportStaff.CancelError = True

    With cdExportSales
      .DialogTitle = "Export to XLS"
      .Filter = "Excel Import File (*.xls)|*.xls"
      .FileName = "Sales" & Format(Now, "YYYYMMDDHHNNSS")
      .ShowSave
    End With
    
    FileExists (cdExportSales.FileName)
    If FExists = True Then
      MsgBox "File Exist!!!"
      btnExportXLSSales_Click
    End If
    
    Set xlApp = Excel.Application
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets.Add
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT id as ItemID, item_desc as Item, quantity_sold as Quantity, total_price as Price,session FROM item_sale WHERE left(submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "' "
    Set rsQuery1 = conn.Execute(SQLText)
    
    totalPrice = 0
    l = 1
    xlSheet.Cells(1, 1) = "Receipt No."
    xlSheet.Cells(1, 2) = "Item"
    xlSheet.Cells(1, 3) = "Quantity"
    xlSheet.Cells(1, 4) = "Price"
    
    Do While Not rsQuery1.EOF
            l = l + 1
            PriceT = rsQuery1("Price")
            xlSheet.Cells(l, 1) = CStr("" & rsQuery1("session") & "")
            xlSheet.Cells(l, 2) = CStr("" & rsQuery1("Item") & "")
            xlSheet.Cells(l, 3) = CStr("" & rsQuery1("Quantity") & "")
            xlSheet.Cells(l, 4) = CStr("RM" & rsQuery1("Price") & "")
            
            totalPrice = totalPrice + PriceT
        rsQuery1.MoveNext
    Loop
    
    xlSheet.Cells(l + 1, 1) = "Total Price"
    xlSheet.Cells(l + 1, 2) = ""
    xlSheet.Cells(l + 1, 3) = ""
    xlSheet.Cells(l + 1, 4) = "RM" & totalPrice
    
    xlBook.SaveAs cdExportSales.FileName
    xlBook.Close
    MsgBox ("Export Success")
End Sub

Private Sub btnFind_Click()

    OpenServer
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where username='" & txtUsernameFind.Text & "' "
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        frmUserProfile.Visible = True
        staff_id2 = rsQuery1("staffID")
        
        'Query for table Staff
        Set rsQuery1 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff where staffID='" & staff_id2 & "' "
        Set rsQuery1 = conn.Execute(SQLText)
             
        If Not rsQuery1.EOF Then
            name2 = rsQuery1("name")
            username = rsQuery1("username")
            email = rsQuery1("email")
            staff_id2 = rsQuery1("staffID")
            
            txtNameAdmin.Text = name2
            txtUsernameAdmin.Text = username
            txtEmailAdmin.Text = email
        Else
            txtNameAdmin.Text = "NULL"
            txtUsernameAdmin.Text = "NULL"
            txtEmailAdmin.Text = "NULL"
        End If
        
        'Query for table Staff_address
        Set rsQuery2 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_address where staffID='" & staff_id2 & "' "
        Set rsQuery2 = conn.Execute(SQLText)
             
        If Not rsQuery2.EOF Then
            address = rsQuery2("address")
            city = rsQuery2("city")
            State = rsQuery2("state")
            zipcode = rsQuery2("zipcode")
            
            txtAddressAdmin.Text = address
            txtCityAdmin.Text = city
            txtStateAdmin.Text = State
            txtZipcodeAdmin.Text = zipcode
        Else
            txtAddressAdmin.Text = "NULL"
            txtCityAdmin.Text = "NULL"
            txtStateAdmin.Text = "NULL"
            txtZipcodeAdmin.Text = "NULL"
        End If
        
        'Query for table Staff_details
        Set rsQuery3 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_details where staffID='" & staff_id2 & "' "
        Set rsQuery3 = conn.Execute(SQLText)
             
        If Not rsQuery3.EOF Then
            tel_home = rsQuery3("tel_home")
            tel_mobile = rsQuery3("tel_mobile")
            position = rsQuery3("position")
            commission = rsQuery3("commission")
            salary_scale = rsQuery3("salary_scale")
            hourly_rate = rsQuery3("hourly_rate")
            status_ = rsQuery3("status")
            
            txtTelHomeAdmin.Text = tel_home
            txtTelMobileAdmin.Text = tel_mobile
            txtPositionAdmin.Text = position
            txtCommissionAdmin.Text = commission
            txtSalaryScale2.Text = salary_scale
            txtHourlyRate2.Text = hourly_rate
            cmbStatusU.Text = status_
        Else
            txtTelHomeAdmin.Text = "NULL"
            txtTelMobileAdmin.Text = "NULL"
            txtPositionAdmin.Text = "NULL"
            txtCommissionAdmin.Text = "NULL"
            txtSalaryScale2.Text = "NULL"
            txtHourlyRate2.Text = "NULL"
            cmbStatusU.Text = "STATUS"
        End If
    Else
        MsgBox "No Record"
        frmUserProfile.Visible = False
    End If
End Sub

Private Sub btnInsert1_Click()
    OpenServer
    'Create table
    Dim rsQuery
    
    Set rsQuery = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where username='" & txtUsernameAdminI.Text & "'"
    Set rsQuery = conn.Execute(SQLText)
         
    If Not rsQuery.EOF Then
        MsgBox "Please use other username."
        txtUsernameAdminI.Text = ""
    ElseIf txtUsernameAdminI.Text = "" And txtNameAdminI.Text = "" And txtPassword.Text = "" And txtEmailAdminI.Text = "" And cmbRole.Text = "ROLE" Then
        MsgBox "Please insert the empty text box."
    Else
        conn.Execute "INSERT INTO staff " _
        & "(name, username,password,role,email) VALUES " _
        & "('" & txtNameAdminI.Text & "','" & txtUsernameAdminI.Text & "',SHA1('" & txtPassword.Text & "'),'" & cmbRole.Text & "','" & txtEmailAdminI.Text & "' )", , adExecuteNoRecords
               
        Set rsQuery = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff where username='" & txtUsernameAdminI.Text & "'"
        Set rsQuery = conn.Execute(SQLText)
        
            If Not rsQuery.EOF Then
                txtStaffIDInsert.Text = rsQuery("staffID")
            End If
        frmUserDetailsInsert.Visible = True
    End If
End Sub

Private Sub btnInsert3_Click()
    OpenServer
    
    If IsNumeric(txtZipcodeAdminI.Text) = True And IsNumeric(txtTel_HomeAdminI.Text) = True And IsNumeric(txtTel_MobileAdminI.Text) = True And IsNumeric(txtCommissionAdminI.Text) = True Then
        conn.Execute "INSERT INTO staff_address(address,city,state,StaffID,zipcode) VALUES ('" & txtAddressAdminI.Text & "' " _
        & " ,'" & txtCityAdminI.Text & "' " _
        & " ,'" & txtStateAdminI.Text & "' " _
        & " ,'" & txtStaffIDInsert.Text & "' " _
        & " ,'" & txtZipcodeAdminI.Text & "')", , adExecuteNoRecords
        
        conn.Execute "INSERT INTO staff_details(tel_home,tel_mobile,position,salary_scale,hourly_rate,StaffID,status,commission) VALUES ('" & Right(txtTel_HomeAdminI.Text, 10) & "' " _
        & " ,'" & Right(txtTel_MobileAdminI.Text, 10) & "' " _
        & " ,'" & txtPositionAdminI.Text & "' " _
        & " ,'" & txtSalaryScale.Text & "' " _
        & " ,'" & txtHourlyRate.Text & "' " _
        & " ,'" & txtStaffIDInsert.Text & "' " _
        & " ,'" & cmbStatusI.Text & "' " _
        & " ,'" & txtCommissionAdminI.Text & "')", , adExecuteNoRecords
    Else
        MsgBox ("Please insert numeric only!!! For Zipcode or Tel Home or Tel Mobile or Commission.")
    End If
    
    frmUserDetailsInsert.Visible = False
    
    txtNameAdminI.Text = ""
    txtUsernameAdminI.Text = ""
    txtPassword.Text = ""
    txtEmailAdminI.Text = ""
    cmbRole.Text = "ROLE"
End Sub

Private Sub btnInventory_Click()
    tabSales.Visible = False
    tabStaff.Visible = False
    tabInventoryAdmin.Visible = True
End Sub

Private Sub btnInventoryStaff_Click()
    tabInventoryStaff.Visible = True
    TabMyProfile.Visible = False
    tabSalesStaff.Visible = False
End Sub

Private Sub btnMonthInv_Click()
    cmbMthInv.Enabled = True
    cmbYrInv.Enabled = True
End Sub

Private Sub btnMonthly_Click()
    cmbMonth.Enabled = True
    cmbYear.Enabled = True
End Sub

Private Sub btnMthSales_Click()
    cmbMthSales.Enabled = True
    cmbYrSales.Enabled = True
End Sub

Private Sub btnNextInvAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx + 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid4.DataSource = rs
    
    txtInvIDAdmin.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
    Set DataGrid4.DataSource = rs2
    
    DataGrid4.Row = Rx
End Sub

Private Sub btnNextInvStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx + 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid3.DataSource = rs
    
    txtInvID.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
    Set DataGrid3.DataSource = rs2
    
    DataGrid3.Row = Rx
End Sub

Private Sub btnNextItemAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx + 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid5.DataSource = rs
    
    txtIIDAdmin.Text = rs!ItemID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
    Set DataGrid5.DataSource = rs2
    
    DataGrid5.Row = Rx
End Sub

Private Sub btnNextStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx + 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid2.DataSource = rs
    
    txtItemIDStaff.Text = rs!ItemID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
    Set DataGrid2.DataSource = rs2
    
    DataGrid2.Row = Rx
End Sub

Private Sub btnPrevInvAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx - 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid4.DataSource = rs
    
    txtInvIDAdmin.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
    Set DataGrid4.DataSource = rs2
    
    DataGrid4.Row = Rx

End Sub

Private Sub btnPrevInvStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx - 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid3.DataSource = rs
    
    txtInvID.Text = rs!ID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ID,item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory", conn
    Set DataGrid3.DataSource = rs2
    
    DataGrid3.Row = Rx
End Sub

Private Sub btnPrevItemAdmin_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx - 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ItemID, item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid5.DataSource = rs
    
    txtIIDAdmin.Text = rs!ItemID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionAdmin.Text & "'", conn
    Set DataGrid5.DataSource = rs2
    
    DataGrid5.Row = Rx
End Sub

Private Sub btnPrevStaff_Click()
    Dim rs, rs2 As ADODB.Recordset
    Dim xCount As Long
    
    On Error Resume Next
    Rx = Rx - 1
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    rs.Open "SELECT id as ItemID, item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
        If rs.EOF = True Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    xCount = rs.RecordCount
        If Rx > rs.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rs.RecordCount - 1
        End If
    rs.Move Rx
    Set DataGrid2.DataSource = rs
    
    txtItemIDStaff.Text = rs!ItemID
    
    rs.Close
    Set rs = Nothing
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.CursorType = adOpenStatic
    rs2.LockType = adLockReadOnly
    rs2.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, price as Price,total_price as Total_Price FROM item_sale where session='" & txtSessionStaff.Text & "'", conn
    Set DataGrid2.DataSource = rs2
    
    DataGrid2.Row = Rx
End Sub

Private Sub btnPrintReceipt_Click()
Dim item_desc As String
Dim price As Double
Dim quantity As Long
Dim totPrice As Double

CommonDialog1.ShowPrinter

    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where session='" & txtSessionAdmin.Text & "'"
    Set rsQuery1 = conn.Execute(SQLText)

    totPrice = 0
    
    Printer.Print "Sales System V1.0 Printing set"
    Printer.Print "#############################################################################"
    Printer.Print vbCrLf
    Printer.Print "Date : " & Now
    Printer.Print "Receipt No. : " & txtSessionAdmin.Text
    Printer.Print vbCrLf
    Printer.Print "Item Description" & vbTab & vbTab & "Quantity Sold" & vbTab & vbTab & "Price"
    Printer.Print vbCrLf

    Do While Not rsQuery1.EOF
        item_desc = rsQuery1("item_desc")
        price = rsQuery1("total_price")
        quantity = rsQuery1("quantity_sold")

        Printer.Print item_desc & vbTab & vbTab & quantity & vbTab & vbTab & "RM " & price
    totPrice = totPrice + price
    rsQuery1.MoveNext
    Loop
    
    Printer.Print vbCrLf
    Printer.Print "Total Price : RM " & totPrice
    Printer.Print "#############################################################################"
    MsgBox ("Print out the report. Please wait....")
    Printer.EndDoc
End Sub

Private Sub btnPrintReceiptStaff_Click()
Dim item_desc As String
Dim price As Double
Dim quantity As Long
Dim totPrice As Double

CommonDialog2.ShowPrinter

    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM item_sale where session=' " & txtSessionStaff.Text & " '"
    Set rsQuery1 = conn.Execute(SQLText)

    totPrice = 0
    
    Printer.Print "Sales System V1.0 Printing set"
    Printer.Print "#############################################################################"
    Printer.Print vbCrLf
    Printer.Print "Date : " & Now
    Printer.Print "Receipt No. : " & txtSessionStaff.Text
    Printer.Print vbCrLf
    Printer.Print "Item Description" & vbTab & vbTab & "Quantity Sold" & vbTab & vbTab & "Price"
    Printer.Print vbCrLf

    Do While Not rsQuery1.EOF
        item_desc = rsQuery1("item_desc")
        price = rsQuery1("total_price")
        quantity = rsQuery1("quantity_sold")

        Printer.Print item_desc & vbTab & vbTab & quantity & vbTab & vbTab & "RM " & price
    totPrice = totPrice + price
    rsQuery1.MoveNext
    Loop
    
    Printer.Print vbCrLf
    Printer.Print "Total Price : RM " & totPrice
    Printer.Print "#############################################################################"
    Printer.EndDoc
    MsgBox ("Print out the report. Please wait....")
End Sub


Private Sub btnPunchIn_Click()

Dim time_now As String

time_now = Format(Now, "YYYY-MM-DD HH:NN:SS")
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_attendance where staffID='" & staff_id & "' and left(time_start,10)='" & Left(time_now, 10) & "'"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        MsgBox "You already punch in today at : " & Format(Now, "HH:NN:SS")
    Else
        conn.Execute "INSERT INTO staff_attendance(time_start,StaffID) VALUES ('" & time_now & "','" & staff_id & "')", , adExecuteNoRecords
        MsgBox "Hi! Your time on starting working at : " & Format(Now, "HH:NN:SS") & " Date : " & Format(Now, "DD-MM-YYYY")
    End If

End Sub

Private Sub btnPunchInAdmin_Click()
Dim time_now As String

time_now = Format(Now, "YYYY-MM-DD HH:NN:SS")
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_attendance where staffID='" & staff_id & "' and left(time_start,10)='" & Left(time_now, 10) & "'"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        MsgBox "You already punch in today at : " & Format(Now, "HH:NN:SS")
    Else
        conn.Execute "INSERT INTO staff_attendance(time_start,StaffID) VALUES ('" & time_now & "','" & staff_id & "')", , adExecuteNoRecords
        MsgBox "Hi! Your time on starting working at : " & Format(Now, "HH:NN:SS") & " Date : " & Format(Now, "DD-MM-YYYY")
    End If
End Sub

Private Sub btnPunchOut_Click(Index As Integer)
Dim time_now As String
Dim hours As String
Dim hours2 As String
Dim time_start As Date
Dim time_end As Date
Dim ended As String
Dim uid As Integer
Dim status As String

time_now = Format(Now, "YYYY-MM-DD HH:NN:SS")

    If MsgBox("Are you sure that you want to Punch Out now?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_attendance where staffID='" & staff_id & "' and left(time_start,10)='" & Left(time_now, 10) & "'"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        uid = rsQuery1("id")
        time_start = rsQuery1("time_start")
        ended = rsQuery1("time_finish")
        ended = Format(ended, "YYYY-MM-DD")
        
        If ended = "1990-01-01" Then
            time_end = Now
            hours = DateDiff("h", time_start, time_end)
            
            Set rsQuery2 = CreateObject("ADODB.Recordset")
            SQLText = "SELECT * FROM staff_details where staffID='" & staff_id & "' "
            Set rsQuery2 = conn.Execute(SQLText)
            
            If Not rsQuery2.EOF Then
                status = rsQuery2("status")
            End If
            
            If status = "Permanent" Then
                hours2 = "1"
            Else
                hours2 = hours
            End If
            
            conn.Execute "UPDATE staff_attendance set time_finish='" & time_now & "',hours_working='" & hours2 & "' where staffID='" & staff_id & "' and id='" & uid & "'", , adExecuteNoRecords
            MsgBox "Goodbye "
            End
        Else
            MsgBox "You are already punch out!!!"
            End
        End If
    Else
        MsgBox "You are not punch in yet!!!"
    End If
End Sub

Private Sub btnPunchOut2_Click(Index As Integer)
Dim time_now As String
Dim hours As String
Dim hours2 As String
Dim status As String
Dim time_start As Date
Dim time_end As Date
Dim ended As String
Dim uid As Integer

time_now = Format(Now, "YYYY-MM-DD HH:NN:SS")

    If MsgBox("Are you sure that you want to Punch Out now?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_attendance where staffID='" & staff_id & "' and left(time_start,10)='" & Left(time_now, 10) & "'"
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        uid = rsQuery1("id")
        time_start = rsQuery1("time_start")
        ended = rsQuery1("time_finish")
        ended = Format(ended, "YYYY-MM-DD")
        
        If ended = "1990-01-01" Then
            time_end = Now
            hours = DateDiff("h", time_start, time_end)
            
            Set rsQuery2 = CreateObject("ADODB.Recordset")
            SQLText = "SELECT * FROM staff_details where staffID='" & staff_id & "' "
            Set rsQuery2 = conn.Execute(SQLText)
            
            If Not rsQuery2.EOF Then
                status = rsQuery2("status")
            End If
            
            If status = "Permanent" Then
                hours2 = "1"
            Else
                hours2 = hours
            End If
            
            conn.Execute "UPDATE staff_attendance set time_finish='" & time_now & "',hours_working='" & hours & "' where staffID='" & staff_id & "' and id='" & uid & "'", , adExecuteNoRecords
            MsgBox "Goodbye "
            End
        Else
            MsgBox "You are already punch out!!!"
            End
        End If
    Else
        MsgBox "You are not punch in yet!!!"
    End If
End Sub

Private Sub btnRetrieve_Click()
    Dim rs1
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    
    lenMonth = Len(cmbMonth.Text)
    lenYear = Len(cmbYear.Text)
    
    If cmbMonth.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYear.Text
    ElseIf cmbMonth.Enabled = True And cmbYear.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYear.Text & "-" & cmbMonth.Text
    End If
    
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.LockType = adLockReadOnly
    rs1.Open "SELECT " _
    & "s.name as Name," _
    & "sd.status as Status," _
    & "((sa.hours_working * sd.salary_scale)/sd.hourly_rate) as Salary, " _
    & "sa.submit_date as Date " _
    & "FROM staff s,staff_details sd, staff_attendance sa WHERE left(sa.submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "' " _
    & " AND sd.staffID=s.staffID and sa.staffID=s.staffID", conn
    Set DataGrid1.DataSource = rs1
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT " _
    & "s.name as Name," _
    & "sd.status as Status," _
    & "((sa.hours_working * sd.salary_scale)/sd.hourly_rate) as Salary, " _
    & "sa.submit_date as Date " _
    & "FROM staff s,staff_details sd, staff_attendance sa WHERE left(sa.submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "' " _
    & " AND sd.staffID=s.staffID and sa.staffID=s.staffID"
    Set rsQuery1 = conn.Execute(SQLText)
            
    totSalary = 0
    
    Do While Not rsQuery1.EOF
        Salary = rsQuery1("Salary")
        totSalary = totSalary + Salary
    rsQuery1.MoveNext
    Loop
    
    lblTotalSalary.Caption = "RM" & totSalary & ".00"
    
End Sub

Private Sub btnSales_Click()
    Dim aSession As String
    Dim aSessionDay As String

    tabSales.Visible = True
    tabStaff.Visible = False
    tabInventoryAdmin.Visible = False
    
    aSessionDay = Format(Now, "DDD")
    aSession = aSessionDay & Format(Now, "DDMMYYHHNNSS")
    
    txtSessionAdmin.Text = aSession
End Sub

Private Sub btnSalesStaff_Click()
    Dim aSession As String
    Dim aSessionDay As String
    
    tabSalesStaff.Visible = True
    TabMyProfile.Visible = False
    tabInventoryStaff.Visible = False
    
    aSessionDay = Format(Now, "DDD")
    aSession = aSessionDay & Format(Now, "DDMMYYHHNNSS")
    
    txtSessionStaff.Text = aSession
        
End Sub

Private Sub btnSearchRS_Click()
    frmSeachReceipt.Show
End Sub

Private Sub btnShowReportInv_Click()
    Dim rs1
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    
    lenMonth = Len(cmbMonth.Text)
    lenYear = Len(cmbYear.Text)
    
    If cmbMthInv.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYrInv.Text
    ElseIf cmbMthInv.Enabled = True And cmbYrInv.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYrInv.Text & "-" & cmbMthInv.Text
    End If
    
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.LockType = adLockReadOnly
    rs1.Open "SELECT item_desc as Item,price,quantity_instock as Quantity,minimum_order_quantity as Min_Quantity,type_of_item as Type_Item  FROM inventory WHERE left(submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "'", conn
    Set DataGrid6.DataSource = rs1
End Sub

Private Sub btnShowReportSales_Click()
    Dim rs1
    Dim lenMonth As Integer
    Dim lenYear As Integer
    Dim TotYear As Integer
    Dim dateString As String
    Dim totSalary As Long
    Dim Salary As Long
    
    lenMonth = Len(cmbMthSales.Text)
    lenYear = Len(cmbYrSales.Text)
    
    If cmbMthSales.Enabled = False Then
        TotYear = lenYear
        dateString = cmbYrSales.Text
    ElseIf cmbMthSales.Enabled = True And cmbYrSales.Enabled = True Then
        TotYear = lenMonth + lenYear + 1
        dateString = cmbYrSales.Text & "-" & cmbMthSales.Text
    End If
    
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.LockType = adLockReadOnly
    rs1.Open "SELECT id as ItemID,item_desc as Item, quantity_sold as Quantity, total_price as Price FROM item_sale where left(submit_date," & TotYear & ")='" & Left(dateString, TotYear) & "'", conn
    Set DataGrid7.DataSource = rs1
End Sub

Private Sub btnStaff_Click()
    tabSales.Visible = False
    tabStaff.Visible = True
    tabInventoryAdmin.Visible = False
End Sub

Private Sub btnUp1Admin_Click()
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where staffID='" & staff_id2 & "' "
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        conn.Execute "UPDATE staff SET name='" & txtNameAdmin.Text & "' " _
        & " ,username='" & txtUsernameAdmin.Text & "' " _
        & " ,email='" & txtEmailAdmin.Text & "' WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
    Else
        MsgBox "Error"
    End If
        MsgBox "Staff " & UCase(txtNameAdmin.Text) & " profile is updated"
        txtNameAdmin.Enabled = False
        txtUsernameAdmin.Enabled = False
        txtEmailAdmin.Enabled = False
        btnUp1Admin.Enabled = False
End Sub

Private Sub btnUp2Admin_Click()
    If IsNumeric(txtZipcodeAdmin.Text) = True Then
        Set rsQuery2 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_address where staffID='" & staff_id2 & "' "
        Set rsQuery2 = conn.Execute(SQLText)
         
        If Not rsQuery2.EOF Then
            conn.Execute "UPDATE staff_address SET address='" & txtAddressAdmin.Text & "' " _
            & " ,city='" & txtCityAdmin.Text & "' " _
            & " ,state='" & txtStateAdmin.Text & "' " _
            & " ,zipcode='" & txtZipcodeAdmin.Text & "' WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
        Else
            conn.Execute "INSERT INTO staff_address(address,city,state,StaffID,zipcode) VALUES ('" & txtAddressAdmin.Text & "' " _
            & " ,'" & txtCityAdmin.Text & "' " _
            & " ,'" & txtStateAdmin.Text & "' " _
            & " ,'" & staff_id2 & "' " _
            & " ,'" & txtZipcodeAdmin.Text & "')", , adExecuteNoRecords
        End If
        MsgBox "Staff " & UCase(txtNameAdmin.Text) & " profile is updated"
        txtAddressAdmin.Enabled = False
        txtCityAdmin.Enabled = False
        txtStateAdmin.Enabled = False
        txtZipcodeAdmin.Enabled = False
        btnUp2Admin.Enabled = False
    Else
        MsgBox ("Please enter only numeric at Zipcode field")
    End If
End Sub

Private Sub btnUp3Admin_Click()
    If IsNumeric(txtCommissionAdmin.Text) = True And IsNumeric(txtTelHomeAdmin.Text) = True And IsNumeric(txtTelMobileAdmin.Text) = True And Len(txtTelHomeAdmin.Text) >= 10 And Len(txtTelMobileAdmin.Text) >= 10 Then
        Set rsQuery3 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_details where staffID='" & staff_id2 & "' "
        Set rsQuery3 = conn.Execute(SQLText)
         
        txtHourlyRate2.Text = "1"
         
        If Not rsQuery3.EOF Then
            conn.Execute "UPDATE staff_details SET tel_home='" & Right(txtTelHomeAdmin.Text, 10) & "' " _
            & " ,tel_mobile='" & Right(txtTelMobileAdmin.Text, 10) & "' " _
            & " ,position='" & txtPositionAdmin.Text & "' " _
            & " ,salary_scale='" & txtSalaryScale2.Text & "' " _
            & " ,hourly_rate='" & txtHourlyRate2.Text & "' " _
            & " ,status='" & cmbStatusU.Text & "' " _
            & " ,commission='" & txtCommissionAdmin.Text & "' WHERE StaffID='" & staff_id2 & "' ", , adExecuteNoRecords
        Else
            conn.Execute "INSERT INTO staff_details(tel_home,tel_mobile,position,salary_scale,hourly_rate,StaffID,status,commission) VALUES ('" & Right(txtTelHomeAdmin.Text, 10) & "' " _
            & " ,'" & Right(txtTelMobileAdmin.Text, 10) & "' " _
            & " ,'" & txtPositionAdmin.Text & "' " _
            & " ,'" & txtSalaryScale2.Text & "' " _
            & " ,'" & txtHourlyRate2.Text & "' " _
            & " ,'" & staff_id2 & "' " _
            & " ,'" & cmbStatusU.Text & "' " _
            & " ,'" & txtCommissionAdmin.Text & "')", , adExecuteNoRecords
        End If
        
        MsgBox "Staff " & UCase(txtNameAdmin.Text) & " profile is updated"
        txtTelHomeAdmin.Enabled = False
        txtTelMobileAdmin.Enabled = False
        txtPositionAdmin.Enabled = False
        txtCommissionAdmin.Enabled = False
        txtSalaryScale2.Enabled = False
        txtHourlyRate2.Enabled = False
        btnUp3Admin.Enabled = False
        cmbStatusU.Enabled = False
    ElseIf Len(txtTelHomeAdmin.Text) < 10 Then
        MsgBox ("Please enter a correct format for Tel No(Home) field")
    ElseIf Len(txtTelMobileAdmin.Text) < 10 Then
        MsgBox ("Please enter a correct format for Tel No(Mobile) field")
    Else
        MsgBox ("Please enter only numeric at Tel No(Home) or Tel No(Mobile) or Commission field")
    End If
End Sub

Private Sub btnUpdateInvAdmin_Click()
    frmUpInventory.Show
End Sub

Private Sub btnUpdateInvStaff_Click()
    frmUpInventory.Show
End Sub

Private Sub btnUpdateP1_Click()
    Dim username2 As String
    Dim username_original As String

    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT username FROM staff"
    Set rsQuery2 = conn.Execute(SQLText)
    
    If Not rsQuery2.EOF Then
        username2 = rsQuery2("username")
    End If
    
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where staffID='" & staff_id & "' "
    Set rsQuery1 = conn.Execute(SQLText)
     
    If Not rsQuery1.EOF Then
        username_original = rsQuery1("username")
        If txtUsername.Text = username2 Then
            MsgBox ("Username already exist. Please change to other username!!!")
            txtUsername.Text = username_original
        Else
            conn.Execute "UPDATE staff SET name='" & txtName.Text & "' " _
            & " ,username='" & txtUsername.Text & "' " _
            & " ,email='" & txtEmail.Text & "' WHERE StaffID='" & staff_id & "' ", , adExecuteNoRecords
            MsgBox "Your profile is updated"
        End If
    Else
        MsgBox "Error"
    End If
End Sub

Private Sub btnUpdateP2_Click()

    If IsNumeric(txtZipcode.Text) = True Then
        Set rsQuery2 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_address where staffID='" & staff_id & "' "
        Set rsQuery2 = conn.Execute(SQLText)
         
        If Not rsQuery2.EOF Then
            conn.Execute "UPDATE staff_address SET address='" & txtAddress.Text & "' " _
            & " ,city='" & txtCity.Text & "' " _
            & " ,state='" & txtState.Text & "' " _
            & " ,zipcode='" & txtZipcode.Text & "' WHERE StaffID='" & staff_id & "' ", , adExecuteNoRecords
        Else
            conn.Execute "INSERT INTO staff_address(address,city,state,StaffID,zipcode) VALUES ('" & txtAddress.Text & "' " _
            & " ,'" & txtCity.Text & "' " _
            & " ,'" & txtState.Text & "' " _
            & " ,'" & staff_id & "' " _
            & " ,'" & txtZipcode.Text & "')", , adExecuteNoRecords
        End If
        
        MsgBox "Your profile is updated"
    Else
        MsgBox ("Please enter only numeric at Zipcode field")
    End If
End Sub

Private Sub btnUpdateP3_Click()

    If IsNumeric(txtTel_Home.Text) = True And IsNumeric(txtTel_Mobile.Text) = True And Len(txtTel_Home.Text) >= 10 And Len(txtTel_Mobile.Text) >= 10 Then
        Set rsQuery3 = CreateObject("ADODB.Recordset")
        SQLText = "SELECT * FROM staff_details where staffID='" & staff_id & "' "
        Set rsQuery3 = conn.Execute(SQLText)
         
        If Not rsQuery3.EOF Then
            conn.Execute "UPDATE staff_details SET tel_home='" & Right(txtTel_Home.Text, 10) & "' " _
            & " ,tel_mobile='" & Right(txtTel_Mobile.Text, 10) & "' " _
            & " ,position='" & txtPosition.Text & "' " _
            & " ,commission='" & txtCommission.Text & "' WHERE StaffID='" & staff_id & "' ", , adExecuteNoRecords
        Else
            conn.Execute "INSERT INTO staff_details(tel_home,tel_mobile,position,StaffID,commission) VALUES ('" & Right(txtTel_Home.Text, 10) & "' " _
            & " ,'" & Right(txtTel_Mobile.Text, 10) & "' " _
            & " ,'" & txtPosition.Text & "' " _
            & " ,'" & staff_id & "' " _
            & " ,'" & txtCommission.Text & "')", , adExecuteNoRecords
        End If
        
        MsgBox "Your profile is updated"
    ElseIf Len(txtTel_Home.Text) < 10 Then
        MsgBox ("Please enter a correct format for Tel No(Home) field")
    ElseIf Len(txtTel_Mobile.Text) < 10 Then
        MsgBox ("Please enter a correct format for Tel No(Mobile) field")
    Else
        MsgBox ("Please enter only numeric at Tel No(Home) or Tel No(Mobile) field")
    End If
End Sub

Private Sub btnYearInv_Click()
    cmbMthInv.Enabled = False
    cmbYrInv.Enabled = True
End Sub

Private Sub btnYearly_Click()
    cmbMonth.Enabled = False
    cmbYear.Enabled = True
End Sub

Private Sub btnYrSales_Click()
    cmbMthSales.Enabled = False
    cmbYrSales.Enabled = True
End Sub

Private Sub cmbStatusU_Click()
    If cmbStatusU.Text = "Permanent" Then
        txtHourlyRate2.Text = "1"
        txtHourlyRate2.Enabled = False
    ElseIf cmbStatusU.Text = "Part Time" Then
        txtHourlyRate2.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    
    If frmMain.txtRole.Text = "Admin" Then
        frmAdmin.Visible = True
    Else
        frmSales.Visible = True
    End If
    
    staff_id = frmMain.txtStaffID.Text
    Rx = 0
    OpenFileConfig
    OpenServer ' Open without ODBC in Control Panel

End Sub

Private Sub btnMyProfile_Click()

    TabMyProfile.Visible = True
    tabSalesStaff.Visible = False
    tabInventoryStaff.Visible = False
    
    OpenServer
    'Create table
    
    'Query for table Staff
    Set rsQuery1 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff where staffID='" & staff_id & "' "
    Set rsQuery1 = conn.Execute(SQLText)
         
    If Not rsQuery1.EOF Then
        name2 = rsQuery1("name")
        username = rsQuery1("username")
        email = rsQuery1("email")
        
        txtName.Text = name2
        txtUsername.Text = username
        txtEmail.Text = email
    Else
        txtName.Text = "NULL"
        txtUsername.Text = "NULL"
        txtEmail.Text = "NULL"
    End If
    
    'Query for table Staff_address
    Set rsQuery2 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_address where staffID='" & staff_id & "' "
    Set rsQuery2 = conn.Execute(SQLText)
         
    If Not rsQuery2.EOF Then
        address = rsQuery2("address")
        city = rsQuery2("city")
        State = rsQuery2("state")
        zipcode = rsQuery2("zipcode")
        
        txtAddress.Text = address
        txtCity.Text = city
        txtState.Text = State
        txtZipcode.Text = zipcode
    Else
        txtAddress.Text = "NULL"
        txtCity.Text = "NULL"
        txtState.Text = "NULL"
        txtZipcode.Text = "NULL"
    End If
    
    'Query for table Staff_details
    Set rsQuery3 = CreateObject("ADODB.Recordset")
    SQLText = "SELECT * FROM staff_details where staffID='" & staff_id & "' "
    Set rsQuery3 = conn.Execute(SQLText)
         
    If Not rsQuery3.EOF Then
        tel_home = rsQuery3("tel_home")
        tel_mobile = rsQuery3("tel_mobile")
        position = rsQuery3("position")
        commission = rsQuery3("commission")
        
        txtTel_Home.Text = tel_home
        txtTel_Mobile.Text = tel_mobile
        txtPosition.Text = position
        txtCommission.Text = commission
    Else
        txtTel_Home.Text = "NULL"
        txtTel_Mobile.Text = "NULL"
        txtPosition.Text = "NULL"
        txtCommission.Text = "NULL"
    End If
End Sub

Private Sub tbSearchRAdmin_Click()
    frmSeachReceipt.Show
End Sub
