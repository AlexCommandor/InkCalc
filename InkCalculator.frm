VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form InkCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ink Coverage Calculator"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "InkCalculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOneHot 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   164
      Text            =   "(in this case results from ALL groups will be included in output file)"
      Top             =   3000
      Width           =   4935
   End
   Begin VB.CheckBox chkUseOneHotfolder 
      Caption         =   "Use ONE input folder for all groups"
      Height          =   255
      Left            =   480
      TabIndex        =   163
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Language/Язык"
      Height          =   855
      Left            =   240
      TabIndex        =   83
      Top             =   6360
      Width           =   2055
      Begin VB.OptionButton optLANG 
         Caption         =   "RUS"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   85
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optLANG 
         Caption         =   "ENG"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   84
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "Double-click on tab to change name!"
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txtGroupsQ 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   81
      Text            =   "Groups quantity:"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cmbNumGroups 
      Height          =   315
      ItemData        =   "InkCalculator.frx":030A
      Left            =   2280
      List            =   "InkCalculator.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   3315
      Width           =   855
   End
   Begin VB.ComboBox cmbRemoveOld 
      Height          =   315
      ItemData        =   "InkCalculator.frx":030E
      Left            =   240
      List            =   "InkCalculator.frx":0327
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "SHEET plate sizes (mm):"
      ForeColor       =   &H00004080&
      Height          =   2415
      Left            =   5280
      TabIndex        =   23
      Top             =   4800
      Width           =   3015
      Begin VB.CommandButton cmdRemoveListSize 
         Caption         =   "Remove checked"
         Height          =   855
         Left            =   1680
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddListSize 
         Caption         =   "Add new"
         Height          =   855
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox lstListSizes 
         Height          =   1860
         ItemData        =   "InkCalculator.frx":0348
         Left            =   240
         List            =   "InkCalculator.frx":034A
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Koef. for SHEET inks (g/m2): "
      ForeColor       =   &H00004080&
      Height          =   2415
      Left            =   2520
      TabIndex        =   12
      Top             =   4800
      Width           =   2775
      Begin VB.TextBox txtInkCoeff 
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   17
         Text            =   "0.90"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtInkCoeff 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   16
         Text            =   "0.90"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtInkCoeff 
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "1.10"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtInkCoeff 
         Height          =   285
         Index           =   3
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   14
         Text            =   "1.00"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtInkCoeff 
         Height          =   285
         Index           =   4
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   13
         Text            =   "2.00"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cyan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   22
         Top             =   735
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Magenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   21
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   20
         Top             =   1455
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Black:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   19
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Other:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   18
         Top             =   1815
         Width           =   615
      End
   End
   Begin VB.CheckBox chkCleanupAfter 
      Caption         =   "Remove PDFs after processing"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   3900
      Width           =   3495
   End
   Begin VB.ComboBox cmbTimeout 
      Height          =   315
      ItemData        =   "InkCalculator.frx":034C
      Left            =   240
      List            =   "InkCalculator.frx":0362
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4950
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5880
      Top             =   7800
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start &processing"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "C:\"
      Top             =   4245
      Width           =   6135
   End
   Begin VB.CommandButton btnSelectOutput 
      Caption         =   "Select &output folder"
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save settings"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdDiscard 
      Cancel          =   -1  'True
      Caption         =   "&Exit programm"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide this window"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Group 1"
      TabPicture(0)   =   "InkCalculator.frx":037F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtInput(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnSelectInput(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Group 2"
      TabPicture(1)   =   "InkCalculator.frx":039B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnSelectInput(1)"
      Tab(1).Control(1)=   "txtInput(1)"
      Tab(1).Control(2)=   "Frame1(1)"
      Tab(1).Control(3)=   "Label1(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Group 3"
      TabPicture(2)   =   "InkCalculator.frx":03B7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(1)=   "btnSelectInput(2)"
      Tab(2).Control(2)=   "txtInput(2)"
      Tab(2).Control(3)=   "Label1(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Group 4"
      TabPicture(3)   =   "InkCalculator.frx":03D3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(1)=   "txtInput(3)"
      Tab(3).Control(2)=   "btnSelectInput(3)"
      Tab(3).Control(3)=   "Label1(3)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Group 5"
      TabPicture(4)   =   "InkCalculator.frx":03EF
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(4)"
      Tab(4).Control(1)=   "txtInput(4)"
      Tab(4).Control(2)=   "btnSelectInput(4)"
      Tab(4).Control(3)=   "Label1(4)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Group 6"
      TabPicture(5)   =   "InkCalculator.frx":040B
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(5)"
      Tab(5).Control(1)=   "txtInput(5)"
      Tab(5).Control(2)=   "btnSelectInput(5)"
      Tab(5).Control(3)=   "Label1(5)"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Group 7"
      TabPicture(6)   =   "InkCalculator.frx":0427
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(6)"
      Tab(6).Control(1)=   "txtInput(6)"
      Tab(6).Control(2)=   "btnSelectInput(6)"
      Tab(6).Control(3)=   "Label1(6)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Group 8"
      TabPicture(7)   =   "InkCalculator.frx":0443
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1(7)"
      Tab(7).Control(1)=   "txtInput(7)"
      Tab(7).Control(2)=   "btnSelectInput(7)"
      Tab(7).Control(3)=   "Label1(7)"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Group 9"
      TabPicture(8)   =   "InkCalculator.frx":045F
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame1(8)"
      Tab(8).Control(1)=   "txtInput(8)"
      Tab(8).Control(2)=   "btnSelectInput(8)"
      Tab(8).Control(3)=   "Label1(8)"
      Tab(8).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   7
         Left            =   -69720
         TabIndex        =   152
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   44
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   157
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   43
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   156
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   42
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   155
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   41
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   154
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   40
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   153
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   44
            Left            =   840
            TabIndex        =   162
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   840
            TabIndex        =   161
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   42
            Left            =   840
            TabIndex        =   160
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   41
            Left            =   600
            TabIndex        =   159
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   40
            Left            =   960
            TabIndex        =   158
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   6
         Left            =   -69720
         TabIndex        =   141
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   39
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   146
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   38
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   145
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   37
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   144
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   36
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   143
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   35
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   142
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   39
            Left            =   840
            TabIndex        =   151
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   840
            TabIndex        =   150
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   37
            Left            =   840
            TabIndex        =   149
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   36
            Left            =   600
            TabIndex        =   148
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   35
            Left            =   960
            TabIndex        =   147
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   5
         Left            =   -69720
         TabIndex        =   130
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   34
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   135
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   33
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   134
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   32
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   133
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   31
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   132
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   30
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   131
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   34
            Left            =   840
            TabIndex        =   140
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   840
            TabIndex        =   139
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   32
            Left            =   840
            TabIndex        =   138
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   31
            Left            =   600
            TabIndex        =   137
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   30
            Left            =   960
            TabIndex        =   136
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   4
         Left            =   -69720
         TabIndex        =   119
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   29
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   124
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   28
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   123
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   27
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   122
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   26
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   121
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   25
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   120
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   29
            Left            =   840
            TabIndex        =   129
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   840
            TabIndex        =   128
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   27
            Left            =   840
            TabIndex        =   127
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   26
            Left            =   600
            TabIndex        =   126
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   25
            Left            =   960
            TabIndex        =   125
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   3
         Left            =   -69720
         TabIndex        =   108
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   24
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   113
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   23
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   112
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   22
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   111
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   21
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   110
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   20
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   109
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   24
            Left            =   840
            TabIndex        =   118
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   840
            TabIndex        =   117
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   22
            Left            =   840
            TabIndex        =   116
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   115
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   20
            Left            =   960
            TabIndex        =   114
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   8
         Left            =   -69720
         TabIndex        =   97
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   45
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   102
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   46
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   101
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   47
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   100
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   48
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   99
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   49
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   98
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   45
            Left            =   960
            TabIndex        =   107
            Top             =   735
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   46
            Left            =   600
            TabIndex        =   106
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   47
            Left            =   840
            TabIndex        =   105
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   840
            TabIndex        =   104
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   49
            Left            =   840
            TabIndex        =   103
            Top             =   1815
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   2
         Left            =   -69720
         TabIndex        =   86
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   19
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   91
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   18
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   90
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   17
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   89
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   16
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   88
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   15
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   87
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   19
            Left            =   840
            TabIndex        =   96
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   840
            TabIndex        =   95
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   94
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   16
            Left            =   600
            TabIndex        =   93
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   15
            Left            =   960
            TabIndex        =   92
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   70
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   2
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   68
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   1
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   1
         Left            =   -69720
         TabIndex        =   56
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   14
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   61
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   13
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   60
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   12
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   59
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   11
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   58
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   10
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   57
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   14
            Left            =   840
            TabIndex        =   66
            Top             =   1815
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   840
            TabIndex        =   65
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   64
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   63
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   10
            Left            =   960
            TabIndex        =   62
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.Frame Frame1 
         Caption         =   "Koeff. for ROLL inks (g/m2): "
         ForeColor       =   &H00008000&
         Height          =   2415
         Index           =   0
         Left            =   5280
         TabIndex        =   43
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   5
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   48
            Text            =   "1.55"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   6
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   47
            Text            =   "1.55"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   7
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   46
            Text            =   "1.95"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   8
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   45
            Text            =   "1.35"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtInkCoeff 
            Height          =   285
            Index           =   9
            Left            =   1575
            MaxLength       =   8
            TabIndex        =   44
            Text            =   "2.00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cyan:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   5
            Left            =   960
            TabIndex        =   53
            Top             =   735
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Magenta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   52
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   840
            TabIndex        =   51
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Black:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   50
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   49
            Top             =   1815
            Width           =   615
         End
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   3
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   3
         Left            =   -74760
         TabIndex        =   41
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   4
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   4
         Left            =   -74760
         TabIndex        =   39
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   5
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   5
         Left            =   -74760
         TabIndex        =   37
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   6
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   6
         Left            =   -74760
         TabIndex        =   35
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   7
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   7
         Left            =   -74760
         TabIndex        =   33
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   8
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "C:\"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.CommandButton btnSelectInput 
         Caption         =   "Select &input folder"
         Height          =   495
         Index           =   8
         Left            =   -74760
         TabIndex        =   31
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   79
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   78
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   77
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   76
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   75
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   74
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   73
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   72
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Watch folder path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   71
         Top             =   840
         Width           =   4815
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "days"
      Height          =   195
      Left            =   1200
      TabIndex        =   29
      Top             =   5820
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Remove records older than"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PDFs MUST BE exported with resolution 100 dpi and with JPEG compression!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout (minutes):"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4710
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Output folder path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   3015
   End
End
Attribute VB_Name = "InkCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tFindResult
   bNoErrors As Boolean
   lResultPosition As Long
   sResultString As String
End Type

Private Type tCVSData
   cvsArray(1 To 28) As String
End Type

Private Enum eCVSData
   a01_Date = 1
   a02_Time = 2
   a03_ZakazNumber = 3
   a04_Name = 4
   a05_SheetNumber = 5
   a06_Tirage = 6
   a07_Black_Face_1 = 7
   a08_Cyan_Face_1 = 8
   a09_Magenta_Face_1 = 9
   a10_Yellow_Face_1 = 10
   a11_Pantone1_Face_1 = 11
   'a12_Pantone2_Face_1 = 12
   a13_Black_Back_1 = 12
   a14_Cyan_Back_1 = 13
   a15_Magenta_Back_1 = 14
   a16_Yellow_Back_1 = 15
   a17_Pantone1_Back_1 = 16
   'a18_Pantone2_Back_1 = 18
   a19_Coverage_Black_Face_1 = 17
   a20_Coverage_Cyan_Face_1 = 18
   a21_Coverage_Magenta_Face_1 = 19
   a22_Coverage_Yellow_Face_1 = 20
   a23_Coverage_Pantone1_Face_1 = 21
   a25_Coverage_Black_Back_1 = 22
   a26_Coverage_Cyan_Back_1 = 23
   a27_Coverage_Magenta_Back_1 = 24
   a28_Coverage_Yellow_Back_1 = 25
   a29_Coverage_Pantone1_Back_1 = 26
   a30_NameOfKoefficientsGroup = 27
   a31_GUID = 28
End Enum

Private Type tSeparationData
   'OrderNumber As String
   'JobName As String
   SheetNumber As String
   SeparationName As String
   SideIsFace As Boolean
   InkWeight As String
   InkArea As String
End Type

Private WithEvents FormSys As FrmSysTray
Attribute FormSys.VB_VarHelpID = -1

Private hndl As Long, CVS_Data As tCVSData
'Private sCVS_

Private Sub CVS_Edit(ByRef aCVS_Array As tCVSData, ByVal aIndex As eCVSData, _
         ByVal sDataString As String)
   aCVS_Array.cvsArray(aIndex) = sDataString
End Sub

Private Sub CVS_Init(ByRef aCVS_Array As tCVSData)
   Dim i As Long
   For i = 1 To UBound(aCVS_Array.cvsArray)
      aCVS_Array.cvsArray(i) = vbNullString
   Next i
End Sub

Private Sub btnSelectInput_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo btnSelectInput_Click_Err
        '</EhHeader>

        Dim sPath As String
        sPath = BrowseForFolder(Me.hWnd)

        If Len(sPath) > 0 Then

            sInputFolder(Index + 1) = sPath
            Me.txtInput(Index).Text = sPath

        End If
         
        If Not WorkFoldersIsNotTheSame Then
            MsgBox sLangStrings(StringIDs.remSomeFoldStillTheSame) & vbCrLf & _
                    sLangStrings(StringIDs.remPleaseSelectProperFolders), _
                    vbExclamation + vbOKOnly, sLangStrings(StringIDs.remIncCalcWarn)
        End If
            
        'Call cmdSave_Click
        '<EhFooter>
        Exit Sub

btnSelectInput_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.btnSelectInput_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub btnSelectOutput_Click()
        '<EhHeader>
        On Error GoTo btnSelectOutput_Click_Err
        '</EhHeader>

        Dim sPath As String
        sPath = BrowseForFolder(Me.hWnd)

        If Len(sPath) > 0 Then

            sOutputFolder = sPath
            Me.txtOutput.Text = sPath

        End If

        If Not WorkFoldersIsNotTheSame Then
            MsgBox sLangStrings(StringIDs.remSomeFoldStillTheSame) & vbCrLf & _
                    sLangStrings(StringIDs.remPleaseSelectProperFolders), _
                    vbExclamation + vbOKOnly, sLangStrings(StringIDs.remIncCalcWarn)
        End If
         'Call cmdSave_Click
        '<EhFooter>
        Exit Sub

btnSelectOutput_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.btnSelectOutput_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Function WorkFoldersIsNotTheSame() As Boolean
    Dim bCheck As Boolean, i As Byte, j As Byte, sTempArray() As String
            
    Call modMain.EnsureHotFoldersExists
    
    If modMain.bUseOneFolder Then
        WorkFoldersIsNotTheSame = True
        Exit Function
    End If
            
    ReDim sTempArray(1 To modMain.numGroups + 1)
    For i = 1 To modMain.numGroups
        sTempArray(i) = txtInput(i - 1).Text
    Next i
    sTempArray(i) = txtOutput.Text

    bCheck = True
    For i = 1 To modMain.numGroups
        For j = i + 1 To modMain.numGroups + 1
            bCheck = bCheck And (sTempArray(i) <> sTempArray(j))
        Next j
    Next i
    
    WorkFoldersIsNotTheSame = bCheck
End Function



Private Sub btnStart_Click()
        'Dim bCheck As Boolean, i As Byte, j As Byte, sTempArray() As String
        Dim i As Byte
        '<EhHeader>
        On Error GoTo btnStart_Click_Err
        '</EhHeader>

        modMain.WriteHotFolders
         'If Me.btnStart.Caption = "Start &processing" Then
         If Me.btnStart.Caption = sLangStrings(StringIDs.remStartProc) Then
         
            
            'ReDim sTempArray(1 To modMain.numGroups + 1)
            'For i = 1 To modMain.numGroups
            '    sTempArray(i) = txtInput(i - 1).Text
            'Next i
            'sTempArray(i) = txtOutput.Text

            'bCheck = True
            'For i = 1 To modMain.numGroups
            '    For j = i + 1 To modMain.numGroups + 1
            '        bCheck = bCheck And (sTempArray(i) <> sTempArray(j))
            '    Next j
            'Next i
            
            'If txtInput.Text = "C:\" Or txtOutput.Text = "C:\" Or txtInput.Text = txtOutput.Text Then
            If Not WorkFoldersIsNotTheSame Then

                Me.Show
                MsgBox sLangStrings(StringIDs.remSomeFoldersTheSame) & vbCrLf & _
                        sLangStrings(StringIDs.remPleaseSelectProperFolders), _
                        vbExclamation + vbOKOnly, sLangStrings(StringIDs.remIncCalcWarn)
                Exit Sub

            End If
            
            'If Not modMain.bUseOneFolder Then
                For i = 1 To modMain.numGroups
                    If Right$(modMain.sInputFolder(i), 2) = ":\" Then
                        Me.Show
                        MsgBox sLangStrings(StringIDs.remSomeFoldersIsRoot) & vbCrLf & _
                                sLangStrings(StringIDs.remItsImpossAllWillDel) & vbCrLf & _
                                sLangStrings(StringIDs.remPleaseSelectProperFolders), _
                                vbError + vbOKOnly, sLangStrings(StringIDs.remIncCalcERR)
                        Exit Sub
                    End If
                Next i
            'End If
                
        
            'Me.btnStart.Caption = "Stop &processing"
            Me.btnStart.Caption = sLangStrings(StringIDs.remStopProc)

            bInProcessing = True
            Me.Hide

            FormSys.TrayIcon = "DEFAULT"
            FormSys.Tooltip = sLangStrings(StringIDs.remWatchFold) ' & sInputFolder
            
            For i = 0 To modMain.numGroups - 1
                Me.btnSelectInput(i).Enabled = False
                Me.Frame1(i).Enabled = False
            Next i
            
            Me.btnSelectOutput.Enabled = False
            Me.cmdSave.Enabled = False
            Me.cmbTimeout.Enabled = False
            Me.chkCleanupAfter.Enabled = False
            'Me.Frame1.Enabled = False
            Me.Frame2.Enabled = False
            Me.Frame3.Enabled = False
            Me.Frame4.Enabled = False
            Me.SSTab1.Enabled = False
            Me.chkUseOneHotfolder.Enabled = False
            Me.cmbRemoveOld.Enabled = False
            Me.cmbNumGroups.Enabled = False
            
            Call WatchFolder

        Else

            Me.Timer1.Enabled = False
            
            Me.SSTab1.Enabled = True
            'Me.btnStart.Caption = "Start &processing"
            Me.btnStart.Caption = sLangStrings(StringIDs.remStartProc)
            
            bInProcessing = False
            'Me.Hide
            For i = 0 To modMain.numGroups - 1
                Me.Frame1(i).Enabled = True
                Me.btnSelectInput(i).Enabled = True
            Next i
            Me.Frame2.Enabled = True
            Me.Frame3.Enabled = True
            Me.Frame4.Enabled = True

            'Me.btnSelectInput.Enabled = True
            Me.btnSelectOutput.Enabled = True
            Me.cmdSave.Enabled = True
            Me.cmbTimeout.Enabled = True
            Me.chkCleanupAfter.Enabled = True
            Me.chkUseOneHotfolder.Enabled = True
            Me.cmbRemoveOld.Enabled = True
            Me.cmbNumGroups.Enabled = True
            
            
            FormSys.TrayIcon = Me
            FormSys.Tooltip = sLangStrings(StringIDs.resInkCalcCaption)

        End If

        '<EhFooter>
        Exit Sub

btnStart_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.btnStart_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub chkCleanupAfter_Click()
    bCleanupAfter = -Abs(Me.chkCleanupAfter.Value)
'    WriteINI "CleanupAfter", CStr(bCleanupAfter)
End Sub

Private Sub chkUseOneHotfolder_Click()
    Dim i As Byte
    modMain.bUseOneFolder = CBool(-Me.chkUseOneHotfolder.Value)
    For i = 1 To modMain.GROUPS_MAX_QUANTITY - 1
        Me.txtInput(i).Enabled = Me.chkUseOneHotfolder.Value - 1
        Me.btnSelectInput(i).Enabled = Me.chkUseOneHotfolder.Value - 1
    Next i
End Sub

Private Sub cmbNumGroups_Click()
    Dim i As Byte, lRes As VbMsgBoxResult
    Me.SSTab1.Tab = 0
    
    'If Val(Me.cmbNumGroups.Text) < modMain.numGroups Then
    '    lRes = MsgBox("You about to remove groups from " & Val(Me.cmbNumGroups.Text) + 1 & " to " & modMain.numGroups & _
    '            vbCrLf & "This cannot be undone! Are you shure to do this?", vbCritical + vbYesNo, "IncCalc warning")
    '    If lRes = vbNo Then
    '        Me.cmbNumGroups.Text = modMain.numGroups
    '        Exit Sub
    '    End If
    'End If
    
    modMain.numGroups = Me.cmbNumGroups.Text
    
    For i = 0 To Me.SSTab1.Tabs - 1
        If i <= modMain.numGroups - 1 Then
            Me.SSTab1.TabEnabled(i) = True
            Me.SSTab1.TabVisible(i) = True
        Else
            Me.SSTab1.TabEnabled(i) = False
            Me.SSTab1.TabVisible(i) = False
        End If
    Next i
    Me.Refresh
End Sub

Private Sub cmbRemoveOld_Change()
   sRemoveOlderThan = Me.cmbRemoveOld.Text
End Sub

Private Sub cmbRemoveOld_Click()
   sRemoveOlderThan = Me.cmbRemoveOld.Text
End Sub

Private Sub cmbRemoveOld_KeyPress(KeyAscii As Integer)
   sRemoveOlderThan = Me.cmbRemoveOld.Text
End Sub

Private Sub cmbTimeout_Change()
    sTimeOut = Me.cmbTimeout.Text
    'WriteINI "TimeOut", sTimeOut
End Sub

Private Sub cmbTimeout_Click()
   sTimeOut = Me.cmbTimeout.Text
   'WriteINI "TimeOut", sTimeOut
End Sub

Private Sub cmbTimeout_KeyPress(KeyAscii As Integer)
   sTimeOut = Me.cmbTimeout.Text
   'WriteINI "TimeOut", sTimeOut
End Sub

Private Sub cmdAbout_Click()
   frmAbout.Show 1
End Sub

Private Sub cmdAddListSize_Click()
   Dim sTmp1 As String, sTmp2 As String, lRes As Long
   sTmp1 = vbNullString
STARTT1:
   Do
      sTmp1 = InputBox(sLangStrings(StringIDs.remEnterWidth), sLangStrings(StringIDs.remInkCalcAddNewPlate))
      If Len(sTmp1) = 0 Then Exit Sub
      If InStr(sTmp1, ".") > 0 Then sTmp1 = vbNullString
      If InStr(sTmp1, ",") > 0 Then sTmp1 = vbNullString
      If InStr(sTmp1, "-") > 0 Then sTmp1 = vbNullString
      If Len(sTmp1) > 0 Then
         If Left$(sTmp1, 1) = "0" Then sTmp1 = vbNullString
      End If
      If Len(sTmp1) = 0 Or Not IsNumeric(sTmp1) Then
         lRes = MsgBox(sLangStrings(StringIDs.remYouAreAboutWrongNumber) & vbCrLf & _
               sLangStrings(StringIDs.remDoYouAbortAddNewSize), vbQuestion + vbYesNo, _
               sLangStrings(StringIDs.remInkCalcAddNewPlate))
         If lRes = vbYes Then sTmp1 = vbNullString: Exit Do Else GoTo STARTT1
      End If
   Loop Until IsNumeric(sTmp1)
   If Len(sTmp1) = 0 Then Exit Sub
   
   sTmp2 = vbNullString
STARTT2:
   Do
      sTmp2 = InputBox(sLangStrings(StringIDs.remEnterHeight), sLangStrings(StringIDs.remInkCalcAddNewPlate))
      If Len(sTmp2) = 0 Then Exit Sub
      If InStr(sTmp2, ".") > 0 Then sTmp2 = vbNullString
      If InStr(sTmp2, ",") > 0 Then sTmp2 = vbNullString
      If InStr(sTmp2, "-") > 0 Then sTmp2 = vbNullString
      If Len(sTmp2) > 0 Then
         If Left$(sTmp2, 1) = "0" Then sTmp2 = vbNullString
      End If
      If Len(sTmp2) = 0 Or Not IsNumeric(sTmp2) Then
         lRes = MsgBox(sLangStrings(StringIDs.remYouAreAboutWrongNumber) & vbCrLf & _
               sLangStrings(StringIDs.remDoYouAbortAddNewSize), vbQuestion + vbYesNo, _
               sLangStrings(StringIDs.remInkCalcAddNewPlate))
         If lRes = vbYes Then sTmp2 = vbNullString: Exit Do Else GoTo STARTT2
      End If
   Loop Until IsNumeric(sTmp2)
   If Len(sTmp2) = 0 Then Exit Sub
   Me.lstListSizes.AddItem sTmp1 & " x " & sTmp2
   sListPlates = Array("")
   ReDim sListPlates(0 To 0)
   For lRes = 0 To Me.lstListSizes.ListCount - 1
      ReDim Preserve sListPlates(0 To lRes)
      sListPlates(lRes) = Me.lstListSizes.List(lRes)
   Next lRes
End Sub

Private Sub cmdDiscard_Click()
        '<EhHeader>
        On Error GoTo cmdDiscard_Click_Err
        '</EhHeader>

        Call FormSys.mExit_Click

        '<EhFooter>
        Exit Sub

cmdDiscard_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.cmdDiscard_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

Private Sub cmdRemoveListSize_Click()
   Dim i As Long
   Do Until Me.lstListSizes.SelCount = 0
      For i = 0 To Me.lstListSizes.ListCount - 1
         If Me.lstListSizes.Selected(i) Then
            Me.lstListSizes.RemoveItem i
            Exit For
         End If
      Next i
   Loop
   If Me.lstListSizes.ListCount > 0 Then
      sListPlates = Array("")
      ReDim sListPlates(0 To 0)
      For i = 0 To Me.lstListSizes.ListCount - 1
         ReDim Preserve sListPlates(0 To i)
         sListPlates(i) = Me.lstListSizes.List(i)
      Next i
   Else
      Set sListPlates = Nothing
   End If
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

        modMain.WriteHotFolders

        If Not WorkFoldersIsNotTheSame Then
            MsgBox sLangStrings(StringIDs.remSomeFoldersTheSame) & vbCrLf & _
                    sLangStrings(StringIDs.remPleaseSelectProperFolders), _
                    vbExclamation + vbOKOnly, sLangStrings(StringIDs.remIncCalcWarn)
        End If


        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.cmdSave_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
    Dim i As Byte ', bFoldEx As Boolean
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        Set FormSys = New FrmSysTray
        Load FormSys
        Set FormSys.FSys = Me
        FormSys.TrayIcon = Me

    For i = 0 To modMain.GROUPS_MAX_QUANTITY - 1
        Me.cmbNumGroups.AddItem CStr(i + 1), i
        Me.btnSelectInput(i).Caption = "Select &input folder N" & CStr(i + 1)
        Me.Label1(i).Caption = "Watch folder path N" & CStr(i + 1) & ":"
        Me.Frame1(i).Caption = "Koef. N" & CStr(i + 1) & " for ROLL inks (g/m2):"
        Me.SSTab1.TabCaption(i) = modMain.sGroupName(i + 1)
    Next i
    
    Me.cmbNumGroups.Text = modMain.numGroups

    For i = 1 To modMain.GROUPS_MAX_QUANTITY
        Me.txtInput(i - 1).Text = sInputFolder(i)
        
        Me.txtInkCoeff(0 + i * 5).Text = sMass_Coeff_ROLL(i)(0)
        Me.txtInkCoeff(1 + i * 5).Text = sMass_Coeff_ROLL(i)(1)
        Me.txtInkCoeff(2 + i * 5).Text = sMass_Coeff_ROLL(i)(2)
        Me.txtInkCoeff(3 + i * 5).Text = sMass_Coeff_ROLL(i)(3)
        Me.txtInkCoeff(4 + i * 5).Text = sMass_Coeff_ROLL(i)(4)
    Next i
    
        Me.txtOutput = sOutputFolder
        Me.cmbTimeout = sTimeOut
        Me.cmbRemoveOld = sRemoveOlderThan
        If IsArray(sListPlates) Then
            Dim ii As Integer
            For ii = 0 To UBound(sListPlates)
               Me.lstListSizes.List(ii) = sListPlates(ii)
            Next ii
        End If
        Me.chkCleanupAfter.Value = Abs(bCleanupAfter)
        
        Me.txtInkCoeff(0).Text = sMass_Coeff_LIST(0)
        Me.txtInkCoeff(1).Text = sMass_Coeff_LIST(1)
        Me.txtInkCoeff(2).Text = sMass_Coeff_LIST(2)
        Me.txtInkCoeff(3).Text = sMass_Coeff_LIST(3)
        Me.txtInkCoeff(4).Text = sMass_Coeff_LIST(4)

        Me.chkUseOneHotfolder.Value = -modMain.bUseOneFolder
        
        Me.optLANG((modMain.lLANG \ 1000) - 1).Value = True
        
        'ChangeLanguage (modMain.lLANG)

        'bFoldEx = True
        'For i = 1 To modMain.numGroups
        '    bFoldEx = bFoldEx And FSO.FolderExists(sInputFolder(i))
        'Next i
        
        ''If FSO.FolderExists(sInputFolder) And FSO.FolderExists(sOutputFolder) Then
        'If bFoldEx And FSO.FolderExists(sOutputFolder) Then

            Call btnStart_Click

        'Else

        '    Me.Show

        'End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.Form_Load " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub


Private Sub WatchFolder()
        '<EhHeader>
        On Error GoTo WatchFolder_Err
        '</EhHeader>
  

           Timer1.Enabled = True
   
        '<EhFooter>
        Exit Sub

WatchFolder_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.WatchFolder " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        'Forcefully terminate any running threads using the END statement
        'Won't work if the thread is busy
            End

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.InkCalc.Form_Unload " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub


Private Sub optLANG_Click(Index As Integer)
    modMain.lLANG = (Index + 1) * 1000
    ChangeLanguage modMain.lLANG, Me
End Sub

Private Sub SSTab1_DblClick()
    Dim sTemp As String
    sTemp = InputBox(sLangStrings(StringIDs.remChNameOfSelGroup), _
            sLangStrings(StringIDs.remInkCalcGrControl), Me.SSTab1.Caption)
    If Len(sTemp) > 0 Then
        Me.SSTab1.Caption = sTemp
        modMain.sGroupName(Me.SSTab1.Tab + 1) = sTemp
    End If
End Sub

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Dim FFO As Object, sTmp As String, i As Integer, iMaxFolders As Integer, j As Integer
    
    
    If Me.chkUseOneHotfolder.Value = 1 Then
        iMaxFolders = 1
    Else
        iMaxFolders = modMain.numGroups
    End If
    
    For i = 1 To iMaxFolders
          
LOOPSS:
        Set FO = FSO.GetFolder(sInputFolder(i))
          
        Me.Timer1.Enabled = False
          
        For Each FFO In FO.SubFolders
    
            sCurrentWorkFolder = FFO.Path
        
                       
            If (FSO.FileExists(sOutputFolder & "\" & FFO.Name & ".txt") = True _
                    And FileLen(sOutputFolder & "\" & FFO.Name & ".txt") = 0) _
                    Or FSO.FileExists(sOutputFolder & "\" & FFO.Name & ".txt") = False Then
                If Not FSO.FolderExists(App.Path & "\" & FFO.Name) Then
                    Err.Clear
                    'On Error GoTo 0
                    FSO.CreateFolder App.Path & "\" & FFO.Name
                             
                    If DoCollecting(sCurrentWorkFolder, sOutputFolder, sTimeOut, i) = True Then
                    
                        'here we need to make some trick - process ONE file numGroups times
                        'so processed file must be deleted only AFTER all
                        If modMain.bUseOneFolder And modMain.numGroups > 1 Then
                            For j = i + 1 To modMain.numGroups
                                DoCollecting sCurrentWorkFolder, sOutputFolder, sTimeOut, j
                            Next j
                        End If
                        
                        FSO.DeleteFolder App.Path & "\" & FFO.Name, True
                        If bCleanupAfter Then
                            FSO.DeleteFile FFO.Path & "\*.pdf", True
                            FFO.Delete True
                        End If
                        Err.Clear
                        'Sleep 2000
                        GoTo LOOPSS
                        ' The line above is bit strange :) It's simply - we spend some time for processing
                        ' (aka Collecting), so in that time period new folders may was added
                             
                    Else
                        FSO.DeleteFolder App.Path & "\" & FFO.Name, True
                        If bCleanupAfter Then
                            FSO.DeleteFile FFO.Path & "\*.pdf", True
                            FFO.Delete True
                        End If
                    End If
                             
                    'Set FFO = Nothing
                    'If bCleanupAfter Then FFO.Delete True: GoTo LOOPSS
                             
                End If
            End If
                        
            DoEvents
        
            If bInProcessing = False Then Me.Timer1.Enabled = False: Exit Sub
            'On Error Resume Next
        Next FFO
        
    Next i
    Me.Timer1.Enabled = True
   
   'Err.Clear
   'On Error GoTo 0

End Sub

Public Function DoCollecting(ByVal sInPath As String, ByVal sOutPath As String, _
                           sTimeOutFuck As String, Optional ByVal iMassKoeffRollIndex As Integer = 1) As Boolean
        '<EhHeader>
        On Error GoTo DoCollecting_Err
        '</EhHeader>

        Dim aTXTFile As String, timeNow As Date, i As Long, j As Long, bRes As Boolean
        Dim nFiles As Integer, iFN As Integer, iFFNN As Integer, sTmp As String
        Dim aFSO As Object, aFO As Object, aFI As Object, sCapt As String
        Dim aTXT_HEADER1 As tCVSData, aTXT_HEADER2 As tCVSData, aTXT_HEADER3 As tCVSData
        Dim sFullSeparationList() As tSeparationData, aJobParamFromFileName As Variant
        Dim aTXT_CURR_CVS As tCVSData, sCleanupData() As String, iSheets() As String
        
        CVS_Edit aTXT_HEADER1, a01_Date, "Date"
        CVS_Edit aTXT_HEADER1, a02_Time, "Time"
        CVS_Edit aTXT_HEADER1, a03_ZakazNumber, "Order number"
        CVS_Edit aTXT_HEADER1, a04_Name, "Job name"
        CVS_Edit aTXT_HEADER1, a05_SheetNumber, "Sheet number"
        CVS_Edit aTXT_HEADER1, a06_Tirage, "Tirage"
        CVS_Edit aTXT_HEADER1, a07_Black_Face_1, "Ink coverage on side, gramm"
        CVS_Edit aTXT_HEADER1, a19_Coverage_Black_Face_1, "Ink area on side, m^2"
        'NEW PARAM!!!
        CVS_Edit aTXT_HEADER1, a30_NameOfKoefficientsGroup, "Koefficients group"
        CVS_Edit aTXT_HEADER1, a31_GUID, "GUID"
        
         
        CVS_Edit aTXT_HEADER2, a07_Black_Face_1, "Face"
        CVS_Edit aTXT_HEADER2, a13_Black_Back_1, "Back"
        CVS_Edit aTXT_HEADER2, a19_Coverage_Black_Face_1, "Face"
        CVS_Edit aTXT_HEADER2, a25_Coverage_Black_Back_1, "Back"
         
        CVS_Edit aTXT_HEADER3, a07_Black_Face_1, "B"
        CVS_Edit aTXT_HEADER3, a08_Cyan_Face_1, "C"
        CVS_Edit aTXT_HEADER3, a09_Magenta_Face_1, "M"
        CVS_Edit aTXT_HEADER3, a10_Yellow_Face_1, "Y"
        CVS_Edit aTXT_HEADER3, a11_Pantone1_Face_1, "P1"
        'CVS_Edit aTXT_HEADER3, a12_Pantone2_Face_1, "P2"

        CVS_Edit aTXT_HEADER3, a13_Black_Back_1, "B"
        CVS_Edit aTXT_HEADER3, a14_Cyan_Back_1, "C"
        CVS_Edit aTXT_HEADER3, a15_Magenta_Back_1, "M"
        CVS_Edit aTXT_HEADER3, a16_Yellow_Back_1, "Y"
        CVS_Edit aTXT_HEADER3, a17_Pantone1_Back_1, "P1"
        'CVS_Edit aTXT_HEADER3, a18_Pantone2_Back_1, "P2"

        CVS_Edit aTXT_HEADER3, a19_Coverage_Black_Face_1, "B"
        CVS_Edit aTXT_HEADER3, a20_Coverage_Cyan_Face_1, "C"
        CVS_Edit aTXT_HEADER3, a21_Coverage_Magenta_Face_1, "M"
        CVS_Edit aTXT_HEADER3, a22_Coverage_Yellow_Face_1, "Y"
        CVS_Edit aTXT_HEADER3, a23_Coverage_Pantone1_Face_1, "P1"
        
        CVS_Edit aTXT_HEADER3, a25_Coverage_Black_Back_1, "B"
        CVS_Edit aTXT_HEADER3, a26_Coverage_Cyan_Back_1, "C"
        CVS_Edit aTXT_HEADER3, a27_Coverage_Magenta_Back_1, "M"
        CVS_Edit aTXT_HEADER3, a28_Coverage_Yellow_Back_1, "Y"
        CVS_Edit aTXT_HEADER3, a29_Coverage_Pantone1_Back_1, "P1"
        
Set aFSO = CreateObject("Scripting.FileSystemObject")
Set aFO = aFSO.GetFolder(sInPath)
iFFNN = FreeFile

Open App.Path & "\" & aFO.Name & "_folder.log" For Append As iFFNN
    
        If Right$(sInPath, Len(sInPath) - InStrRev(sInPath, "\")) = "New folder" Or _
           Right$(sInPath, Len(sInPath) - InStrRev(sInPath, "\")) = "Новая папка" Then
            
            DoCollecting = False
            Exit Function

        End If
    
        On Error Resume Next
    
'108     Set aFSO = CreateObject("Scripting.FileSystemObject")
'110     Set aFO = aFSO.GetFolder(sInPath)


        If Err.Number <> 0 Then _
            Print #iFFNN, Time() & Chr$(9) & "Start checking for files in " & sInPath & _
                " failed! Folder is missing???": DoCollecting = False: Exit Function 'Folder missing?????????????
        
        timeNow = Time()

    Print #iFFNN, Time() & Chr$(9) & "Start checking for files in " & sInPath


            
      'aTXTFile = sOutPath & "\" & aFO.Name & ".txt"
      
      'If aFSO.FileExists(aTXTFile) Then aFSO.DeleteFile aTXTFile, True

         If aFO.Files.Count = 0 Then 'no files in folder!
            Print #iFFNN, Time() & Chr$(9) & "No files found in folder! Exiting..."
            Close #iFFNN
            Set aFI = Nothing
            Set aFO = Nothing
            Set aFSO = Nothing
            DoCollecting = False
            Exit Function
         End If

         iFN = FreeFile()

         
         
'123      Open aTXTFile For Append As iFN
sCapt = FormSys.Tooltip

         For Each aFI In aFO.Files

FormSys.Tooltip = sLangStrings(StringIDs.remAnalFile) & aFI.Path
FormSys.TrayIcon = "PROCESSING"

            If bInProcessing = False Then Me.Timer1.Enabled = False: DoCollecting = False: Exit Function
            
    Print #iFFNN, Time() & Chr$(9) & "Analizing " & aFI.Path & " file."

            'If CalculateInks(aFI.Path, iFFNN, aFO.Name) = True Then
            If CalculateInks(aFI.Path, iFFNN, iMassKoeffRollIndex) = True Then
               
    Print #iFFNN, Time() & Chr$(9) & "CalculateInks for " & aFI.Path & " proceeded. Saving output..."

               'Open App.Path & "\" & aFO.Name & "_process.txt" For Append As iFN
               iFN = FreeFile()
               Open sResultFile For Input As iFN
                  Line Input #iFN, sTmp
               Close #iFN
               If sTmp <> Join(aTXT_HEADER1.cvsArray, vbTab) Then 'Wrong file header! We need to create new file
LOOPS1:
                  Err.Clear
                  iFN = FreeFile()
                  Open sResultFile For Output As iFN
                     If Err.Number <> 0 Then
                        MsgBox sLangStrings(StringIDs.remUnableWriteFile) & sResultFile & _
                        sLangStrings(StringIDs.remMaybeItLockCloseAndTry), _
                        vbCritical, sLangStrings(StringIDs.resInkCalcCaption)
                        GoTo LOOPS1
                     End If
                     Print #iFN, Join(aTXT_HEADER1.cvsArray, vbTab)
                     Print #iFN, Join(aTXT_HEADER2.cvsArray, vbTab)
                     Print #iFN, Join(aTXT_HEADER3.cvsArray, vbTab)
                  Close #iFN
LOOPS2:
                  Err.Clear
                  Open sResultFile & ".txt" For Output As iFN
                     If Err.Number <> 0 Then
                        MsgBox sLangStrings(StringIDs.remUnableWriteFile) & sResultFile & ".txt" & _
                        sLangStrings(StringIDs.remMaybeItLockCloseAndTry), _
                        vbCritical, sLangStrings(StringIDs.resInkCalcCaption)
                        GoTo LOOPS2
                     End If
                     Print #iFN, Join(aTXT_HEADER1.cvsArray, vbTab)
                     Print #iFN, Join(aTXT_HEADER2.cvsArray, vbTab)
                     Print #iFN, Join(aTXT_HEADER3.cvsArray, vbTab)
                  Close #iFN
               End If
               
               'Cleanup CVS file from sRemoveOlderThan days old strings
               iFN = FreeFile()
               i = 1
               ReDim sCleanupData(1 To 1)
               Open sResultFile & ".txt" For Input As iFN
                  Do While Not EOF(iFN)
                     ReDim Preserve sCleanupData(1 To i)
                     Line Input #iFN, sCleanupData(i)
                     i = i + 1
                  Loop
               Close #iFN
               
               iFN = FreeFile()
               Open sResultFile For Output As iFN
                  Print #iFN, Join(aTXT_HEADER1.cvsArray, vbTab)
                  Print #iFN, Join(aTXT_HEADER2.cvsArray, vbTab)
                  Print #iFN, Join(aTXT_HEADER3.cvsArray, vbTab)
                  For i = 4 To UBound(sCleanupData)
                     If Date - CDate(Split(sCleanupData(i), vbTab)(0)) < Val(sRemoveOlderThan) Then
                        Print #iFN, sCleanupData(i)
                     End If
                  Next i
               Close #iFN
               Open sResultFile & ".txt" For Output As iFN
                  Print #iFN, Join(aTXT_HEADER1.cvsArray, vbTab)
                  Print #iFN, Join(aTXT_HEADER2.cvsArray, vbTab)
                  Print #iFN, Join(aTXT_HEADER3.cvsArray, vbTab)
                  For i = 4 To UBound(sCleanupData)
                     If Date - CDate(Split(sCleanupData(i), vbTab)(0)) < Val(sRemoveOlderThan) Then
                        Print #iFN, sCleanupData(i)
                     End If
                  Next i
               Close #iFN
               
               aJobParamFromFileName = Split(Replace$(aFI.Name, ".pdf", vbNullString), "~")
               
               'here we must build list of sheets numbers founded during processing
               ReDim iSheets(1 To 1)
               iSheets(1) = Split(sPages(1), vbTab)(1)
               For i = 1 To UBound(sPages)
                  bRes = True
                  For j = 1 To UBound(iSheets)
                     If iSheets(j) = Split(sPages(i), vbTab)(1) Then bRes = False: Exit For
                  Next j
                  If bRes Then
                     ReDim Preserve iSheets(1 To UBound(iSheets) + 1)
                     iSheets(UBound(iSheets)) = Split(sPages(i), vbTab)(1)
                  End If
               Next i
               
               ReDim sFullSeparationList(1 To UBound(iSheets) * 10)
               For i = 0 To UBound(iSheets) * 2 - 1
                 sFullSeparationList(1 + i * 5).SeparationName = "BLACK"
                 sFullSeparationList(1 + i * 5).SheetNumber = iSheets(i \ 2 + 1)
                 sFullSeparationList(1 + i * 5).SideIsFace = (i Mod 2) - 1
                 sFullSeparationList(1 + i * 5).InkWeight = "0.000000"
                 sFullSeparationList(1 + i * 5).InkArea = "0.000000"
                 
                 sFullSeparationList(2 + i * 5).SeparationName = "CYAN"
                 sFullSeparationList(2 + i * 5).SheetNumber = iSheets(i \ 2 + 1)
                 sFullSeparationList(2 + i * 5).SideIsFace = (i Mod 2) - 1
                 sFullSeparationList(2 + i * 5).InkWeight = "0.000000"
                 sFullSeparationList(2 + i * 5).InkArea = "0.000000"
                 
                 sFullSeparationList(3 + i * 5).SeparationName = "MAGENTA"
                 sFullSeparationList(3 + i * 5).SheetNumber = iSheets(i \ 2 + 1)
                 sFullSeparationList(3 + i * 5).SideIsFace = (i Mod 2) - 1
                 sFullSeparationList(3 + i * 5).InkWeight = "0.000000"
                 sFullSeparationList(3 + i * 5).InkArea = "0.000000"
                 
                 sFullSeparationList(4 + i * 5).SeparationName = "YELLOW"
                 sFullSeparationList(4 + i * 5).SheetNumber = iSheets(i \ 2 + 1)
                 sFullSeparationList(4 + i * 5).SideIsFace = (i Mod 2) - 1
                 sFullSeparationList(4 + i * 5).InkWeight = "0.000000"
                 sFullSeparationList(4 + i * 5).InkArea = "0.000000"
                 
                 sFullSeparationList(5 + i * 5).SeparationName = "P1"
                 sFullSeparationList(5 + i * 5).SheetNumber = iSheets(i \ 2 + 1)
                 sFullSeparationList(5 + i * 5).SideIsFace = (i Mod 2) - 1
                 sFullSeparationList(5 + i * 5).InkWeight = "0.000000"
                 sFullSeparationList(5 + i * 5).InkArea = "0.000000"
               Next i
               
               'now we must check plates list on order of separations
               For i = 1 To UBound(sFullSeparationList)
                  For j = 1 To UBound(sPages)
                     If Split(UCase$(sPages(j)), vbTab)(0) = sFullSeparationList(i).SeparationName And _
                        Split(UCase$(sPages(j)), vbTab)(1) = sFullSeparationList(i).SheetNumber Then
                           If (Split(UCase$(sPages(j)), vbTab)(2) = "FACE" And sFullSeparationList(i).SideIsFace = True) Or _
                              (Split(UCase$(sPages(j)), vbTab)(2) = "BACK" And sFullSeparationList(i).SideIsFace = False) Then
                                    sFullSeparationList(i).InkWeight = Split(UCase$(sPages(j)), vbTab)(3)
                                    sFullSeparationList(i).InkArea = Split(UCase$(sPages(j)), vbTab)(4)
                                    Exit For
                           End If
                     End If
                  Next j
               Next i
               
               For i = 1 To UBound(iSheets)
                  CVS_Init aTXT_CURR_CVS
                  CVS_Edit aTXT_CURR_CVS, a01_Date, CStr(Date)
                  CVS_Edit aTXT_CURR_CVS, a02_Time, CStr(Time)
                  If IsArray(aJobParamFromFileName) Then
                     CVS_Edit aTXT_CURR_CVS, a03_ZakazNumber, aJobParamFromFileName(0)
                     CVS_Edit aTXT_CURR_CVS, a04_Name, aJobParamFromFileName(1)
                  Else
                     CVS_Edit aTXT_CURR_CVS, a03_ZakazNumber, aFO.Name
                     CVS_Edit aTXT_CURR_CVS, a04_Name, aFO.Name
                  End If
                  CVS_Edit aTXT_CURR_CVS, a05_SheetNumber, iSheets(i)
                  
                  CVS_Edit aTXT_CURR_CVS, a07_Black_Face_1, sFullSeparationList((i - 1) * 10 + 1).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a08_Cyan_Face_1, sFullSeparationList((i - 1) * 10 + 2).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a09_Magenta_Face_1, sFullSeparationList((i - 1) * 10 + 3).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a10_Yellow_Face_1, sFullSeparationList((i - 1) * 10 + 4).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a11_Pantone1_Face_1, sFullSeparationList((i - 1) * 10 + 5).InkWeight
                  
                  CVS_Edit aTXT_CURR_CVS, a13_Black_Back_1, sFullSeparationList((i - 1) * 10 + 6).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a14_Cyan_Back_1, sFullSeparationList((i - 1) * 10 + 7).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a15_Magenta_Back_1, sFullSeparationList((i - 1) * 10 + 8).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a16_Yellow_Back_1, sFullSeparationList((i - 1) * 10 + 9).InkWeight
                  CVS_Edit aTXT_CURR_CVS, a17_Pantone1_Back_1, sFullSeparationList((i - 1) * 10 + 10).InkWeight
                  
                  CVS_Edit aTXT_CURR_CVS, a19_Coverage_Black_Face_1, sFullSeparationList((i - 1) * 10 + 1).InkArea
                  CVS_Edit aTXT_CURR_CVS, a20_Coverage_Cyan_Face_1, sFullSeparationList((i - 1) * 10 + 2).InkArea
                  CVS_Edit aTXT_CURR_CVS, a21_Coverage_Magenta_Face_1, sFullSeparationList((i - 1) * 10 + 3).InkArea
                  CVS_Edit aTXT_CURR_CVS, a22_Coverage_Yellow_Face_1, sFullSeparationList((i - 1) * 10 + 4).InkArea
                  CVS_Edit aTXT_CURR_CVS, a23_Coverage_Pantone1_Face_1, sFullSeparationList((i - 1) * 10 + 5).InkArea
                  
                  CVS_Edit aTXT_CURR_CVS, a25_Coverage_Black_Back_1, sFullSeparationList((i - 1) * 10 + 6).InkArea
                  CVS_Edit aTXT_CURR_CVS, a26_Coverage_Cyan_Back_1, sFullSeparationList((i - 1) * 10 + 7).InkArea
                  CVS_Edit aTXT_CURR_CVS, a27_Coverage_Magenta_Back_1, sFullSeparationList((i - 1) * 10 + 8).InkArea
                  CVS_Edit aTXT_CURR_CVS, a28_Coverage_Yellow_Back_1, sFullSeparationList((i - 1) * 10 + 9).InkArea
                  CVS_Edit aTXT_CURR_CVS, a29_Coverage_Pantone1_Back_1, sFullSeparationList((i - 1) * 10 + 10).InkArea
                  
                  CVS_Edit aTXT_CURR_CVS, a30_NameOfKoefficientsGroup, modMain.sGroupName(iMassKoeffRollIndex)
                  CVS_Edit aTXT_CURR_CVS, a31_GUID, modMain.GenerateRandomGUID
                                                
                  iFN = FreeFile()
                  Open sResultFile For Append As iFN
                     Print #iFN, Replace$(Join(aTXT_CURR_CVS.cvsArray, vbTab), ",", ".")
                  Close #iFN
   
                  Open sResultFile & ".txt" For Append As iFN
                     Print #iFN, Replace$(Join(aTXT_CURR_CVS.cvsArray, vbTab), ",", ".")
                  Close #iFN
               Next i

'               Open sResultFile For Append As iFN
'                  Print #iFN, vbCrLf & vbCrLf & Date & vbTab & Time() & vbCrLf & vbCrLf & aFO.Name & vbCrLf
'                  Print #iFN, "SEPARATION" & vbTab & "PAGE" & vbTab & "INK, MxM" & vbTab & "WEIGHT, G" & vbCrLf
'                  Print #iFN, Join(sPages, vbCrLf)
'               Close #iFN
               
'               Open sResultFile & ".xls" For Append As iFN
'                  Print #iFN, vbCrLf & vbCrLf & Date & vbTab & Time() & vbCrLf & vbCrLf & aFO.Name & vbCrLf
'                  Print #iFN, "SEPARATION" & vbTab & "PAGE" & vbTab & "INK, MxM" & vbTab & "WEIGHT, G" & vbCrLf
'                  Print #iFN, Join(sPages, vbCrLf)
'               Close #iFN
               
    Print #iFFNN, Time() & Chr$(9) & "Saved data into " & aTXTFile & "."

               aFSO.DeleteFolder App.Path & "\" & aFO.Name, True

    Print #iFFNN, Time() & Chr$(9) & "Deleted folder " & App.Path & "\" & aFO.Name & "."
            
            Else
               
         Print #iFFNN, Time() & Chr$(9) & "CalculateInks in " & aFI.Path & " failed!"
               Set aFI = Nothing
               Set aFO = Nothing
               Set aFSO = Nothing
               Close #iFFNN
               DoCollecting = False
               FormSys.Tooltip = sCapt
               FormSys.TrayIcon = "DEFAULT"
               Exit Function
               
            End If
            
            
            'Operation timeout !!!
            If Minute(Time() - timeNow) > Val(sTimeOutFuck) Then
                Print #iFFNN, Time() & Chr$(9) & "Timeout waiting for files!!! Aborting ...   wait: " _
                    & Minute(Time() - timeNow) & ", timeout: " & sTimeOutFuck
               DoCollecting = False
               Exit Function
            End If
            
FormSys.Tooltip = sCapt
FormSys.TrayIcon = "DEFAULT"
            
            Sleep 2000
            
      Next
      Close #iFN
      
'Err.Clear
'Print #iFFNN, Time() & Chr$(9) & "Output file size is  " & FileLen(App.Path & "\" & aFO.Name & "_process.txt") & " bytes."

 '     If FileLen(App.Path & "\" & aFO.Name & "_process.txt") > 10 Then
 '        aFSO.CopyFile App.Path & "\" & aFO.Name & "_process.txt", aTXTFile, True
 '     End If
 '     aFSO.DeleteFile App.Path & "\" & aFO.Name & "_process.txt", True

'Print #iFFNN, Time() & Chr$(9) & "Processed " & nFiles & "files."

      
        Set aFI = Nothing
        Set aFO = Nothing
'        aFSO.DeleteFolder sInPath, True
        Set aFSO = Nothing
    
Print #iFFNN, Time() & Chr$(9) & "Pizdets! :)"
      Close #iFFNN
         If Err.Number <> 0 Then DoCollecting = False Else DoCollecting = True
        '<EhFooter>
        Exit Function

DoCollecting_Err:

    Print #iFFNN, Time() & Chr$(9) & Err.Description & " at line " & Erl
    Close #iFFNN
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.DoCollecting " & _
               "at line " & Erl
        End
        '</EhFooter>
End Function

'Public Function CalculateInks(ByVal sFile As String, Optional ByVal iLogFileNum As Integer, _
         Optional ByVal sWorkFolder As String = vbNullString) As Boolean
   
Public Function CalculateInks(ByVal sFile As String, Optional ByVal iLogFileNum As Integer, _
                Optional ByVal iMassKoeffRollIndex As Integer = 1) As Boolean
   'Const pcxKoeff As Double = 2.83464567
   Dim MyObj As GflAx.GflAx, W As Currency, H As Currency, i As Currency, j As Currency
   Dim SUM As Currency, FULL As Currency, DPI As Integer
   Dim Wcm As Currency, Hcm As Currency, FULLcm2 As Currency
   Dim tmp As Currency, FFSO As Object, bFileReadError As Boolean
   Dim bBuff() As Byte, WSH As Object, bGetLogPages As Boolean, iIFN As Integer
   Dim dWeight As Single, arListSize As Variant, isListSize As Boolean
   
   On Error Resume Next
   
   CalculateInks = False
   
   DPI = 100
   
   Set MyObj = New GflAx.GflAx
   Set WSH = CreateObject("WScript.Shell")
   Set FFSO = CreateObject("Scripting.FileSystemObject")
   
   Print #iLogFileNum, Time() & Chr$(9) & "Trying to get logical page numbering from " & sFile & "."
   
   bFileReadError = False
   bGetLogPages = GetLogicalPages(sFile, bFileReadError)
   If bGetLogPages = False And bFileReadError = False Then
      Print #iLogFileNum, Time() & Chr$(9) & "Getting logical page numbering from " & sFile & " failed!"
      CalculateInks = False
      Exit Function
   ElseIf bGetLogPages = False And bFileReadError = True Then
      Print #iLogFileNum, Time() & Chr$(9) & "Getting logical page numbering from " & sFile & " failed - file read error! (timeout or other I/O error)"
      CalculateInks = False
      Exit Function
   End If
   
   Print #iLogFileNum, Time() & Chr$(9) & "Getting logical page numbering from " & sFile & " with no errors."
   Print #iLogFileNum, Time() & Chr$(9) & "Found " & UBound(sPages) & " separated pages"
   
   
   
  For i = LBound(sPages) To UBound(sPages)
  
   Print #iLogFileNum, Time() & Chr$(9) & "Run command for convert pdf to image..."
   'Print #iLogFileNum, Time() & Chr$(9) & Chr$(34) & App.Path & "\pdftoppm.exe" & Chr$(34) & " -aa no -aaVector no -gray -f " & _
         CStr(i) & " -l " & CStr(i) & " -r " & _
         CStr(DPI) & " -q " & Chr$(34) & sFile & Chr$(34) & " " & _
         Chr$(34) & sFile & "-page" & Chr$(34)
         
'   If WSH.Run(Chr$(34) & App.Path & "\pdftoppm.exe" & Chr$(34) & " -aa no -aaVector no -gray -f " & _
         CStr(i) & " -l " & CStr(i) & " -r " & _
         CStr(DPI) & " -q " & Chr$(34) & sFile & Chr$(34) & " " & _
         Chr$(34) & sFile & "-page" & Chr$(34), 0&, True) <> 0 Then
         
   If WSH.Run(Chr$(34) & App.Path & "\pdfimages.exe" & Chr$(34) & " -f " & CStr(i) & " -l " & CStr(i) & _
         " -j -q " & Chr$(34) & sFile & Chr$(34) & " " & _
         Chr$(34) & sFile & "-page" & Chr$(34), 0&, True) <> 0 Then
         
      Print #iLogFileNum, Time() & Chr$(9) & "Failed to convert page # " & i & " to image!"
      GoTo NNEXTT

   End If
   
   'MyObj.LoadBitmap sFile & "-page-" & Format$(i, "000000") & ".pgm"
   MyObj.LoadBitmap sFile & "-page-000.jpg"
   
   W = MyObj.Width
   H = MyObj.Height

   If W = 0 Or H = 0 Then
      Print #iLogFileNum, Time() & Chr$(9) & "Failed to load image for page # " & i
      GoTo NNEXTT
   End If

   FULL = W * H
   Wcm = W * 2.54 / DPI
   Hcm = H * 2.54 / DPI
   FULLcm2 = Wcm * Hcm
   
   Print #iLogFileNum, Time() & Chr$(9) & "Image boundaries for page # " & i & " - " & _
            Format$(Wcm, "0.00") & " x " & Format$(Hcm, "0.00") & " cm"
   
   'MyObj.Negative
   'MyObj.ChangeColorDepth &H100, 0, 1
   MyObj.SaveFormat = AX_PGM
   MyObj.SaveBitmap sFile & "-page-000.pgm"

   'Erase bBuff
   
   iIFN = FreeFile()
   Open sFile & "-page-000" & ".pgm" For Binary Access Read As #iIFN
      ReDim bBuff(0 To LOF(iIFN) - 1)
      Get #iIFN, , bBuff
   Close #iIFN
   'bBuff = MyObj.SendBinary
   
   'FFSO.DeleteFile sFile & "-page-" & Format$(i, "000000") & ".pgm", True
   FFSO.DeleteFile sFile & "-page-000.jpg", True
   FFSO.DeleteFile sFile & "-page-000.pgm", True
   
   SUM = 0
   
   For j = 0 To UBound(bBuff)
         'SUM = SUM + 255 - bBuff(j)
         SUM = SUM + bBuff(j)
   Next j
   
   FULL = FULL * 255
   
   isListSize = False
   If IsArray(sListPlates) Then ' we have some LIST plates defined, so we check current size for LIST
      For j = LBound(sListPlates) To UBound(sListPlates)
         arListSize = Split(sListPlates(j), " x ")
         ' if difference between curr sizes and current LIST plates is <= 1 cm then we have LIST plate
         If (Abs((arListSize(0) - Wcm * 10)) <= 1 And (Abs(arListSize(1) - Hcm * 10)) <= 1) Or _
               (Abs((arListSize(1) - Wcm * 10)) <= 1 Or (Abs(arListSize(0) - Hcm * 10)) <= 1) Then
            isListSize = True
         End If
      Next j
   End If
   
   dWeight = FULLcm2 * SUM / (FULL * 10000) ' Вес в граммах на лист
   If InStr(1, UCase$(sPages(i)), "CYAN") = 1 Then
      If isListSize Then
         dWeight = dWeight * Val(sMass_Coeff_LIST(0))
      Else
         dWeight = dWeight * Val(sMass_Coeff_ROLL(iMassKoeffRollIndex)(0))
      End If
   ElseIf InStr(1, UCase$(sPages(i)), "MAGENTA") = 1 Then
      If isListSize Then
         dWeight = dWeight * Val(sMass_Coeff_LIST(1))
      Else
         dWeight = dWeight * Val(sMass_Coeff_ROLL(iMassKoeffRollIndex)(1))
      End If
   ElseIf InStr(1, UCase$(sPages(i)), "YELLOW") = 1 Then
      If isListSize Then
         dWeight = dWeight * Val(sMass_Coeff_LIST(2))
      Else
         dWeight = dWeight * Val(sMass_Coeff_ROLL(iMassKoeffRollIndex)(2))
      End If
   ElseIf InStr(1, UCase$(sPages(i)), "BLACK") = 1 Then
      If isListSize Then
         dWeight = dWeight * Val(sMass_Coeff_LIST(3))
      Else
         dWeight = dWeight * Val(sMass_Coeff_ROLL(iMassKoeffRollIndex)(3))
      End If
   Else
      If isListSize Then
         dWeight = dWeight * Val(sMass_Coeff_LIST(4))
      Else
         dWeight = dWeight * Val(sMass_Coeff_ROLL(iMassKoeffRollIndex)(4))
      End If
   End If

'   If Len(sWorkFolder) > 0 Then
'      sPages(i) = sPages(i) & vbTab & Format$(FULLcm2 * SUM / (FULL * 10000), "0.000000") & _
            vbTab & Format$(dWeight, "0.000000") & vbCrLf
'   Else
'      sPages(i) = sPages(i) & vbTab & Format$(FULLcm2 * SUM / (FULL * 10000), "0.000000") & _
            vbTab & Format$(dWeight, "0.000000") & vbCrLf
'   End If
   sPages(i) = sPages(i) & vbTab & Format$(dWeight, "0.000000") & vbTab & Format$(FULLcm2 * SUM / (FULL * 10000), "0.000000")
   Print #iLogFileNum, Time() & Chr$(9) & "Data for page # " & i & " = " & vbCrLf & sPages(i)
NNEXTT:
   If bInProcessing = False Then Me.Timer1.Enabled = False: CalculateInks = False: Exit Function
   DoEvents
  Next i
  Set MyObj = Nothing
  Set WSH = Nothing
  Set FFSO = Nothing
  
  If Err.Number = 0 Then CalculateInks = True Else Err.Clear: CalculateInks = False
  On Error GoTo 0
End Function

Private Function GetLogicalPages(ByVal sFile As String, _
               Optional ByRef bErrorInFile As Boolean = False) As Boolean
   Dim sTmp As String, vArr() As String, i As Long
   Dim bBBuf() As Byte, iFnumer As Integer, vFuck As Variant, bZeroLenghtCheck As Boolean
   Dim tStartTime As Date, FindResult() As tFindResult, FindResult2() As tFindResult
   Dim aSplitted As Variant, lSheetNumber As Long
   
   On Error Resume Next
   
   ReDim bBBuf(1 To 10)
   iFnumer = FreeFile
   bZeroLenghtCheck = False
STILL:
   If FileLen(sFile) < 10 Then
      Err.Clear
      If bZeroLenghtCheck = False Then
         bZeroLenghtCheck = True
         Sleep 5000
         GoTo STILL
      Else
         Err.Clear
         On Error GoTo 0
         bErrorInFile = True
         GetLogicalPages = False
         Exit Function
      End If
   End If
   tStartTime = Now()
   Open sFile For Binary Access Read As iFnumer
      If Err.Number <> 0 Then ' file still busy! (maybe still writing)
         Close #iFnumer
         Err.Clear
         Sleep 3000
         GoTo STILL
      End If
      Get #iFnumer, LOF(iFnumer) - 9, bBBuf
   Close #iFnumer
   sTmp = vbNullString
   For i = 1 To 10
      sTmp = sTmp & Chr$(bBBuf(i))
   Next i
   If Not (sTmp Like "*%EOF*") Then
      ' file not locked but still doesnt have EOF marker! fucking shit
      If Minute(Now() - tStartTime) >= Val(sTimeOut) Then ' fixed timeout for waiting for PDF!
         Err.Clear
         bErrorInFile = True
         GetLogicalPages = False
      End If
      Sleep 3000
      GoTo STILL
   End If
   
'   vArr = Split(Replace$(StrConv(bBBuf, vbUnicode), vbCr, vbLf), vbLf)
'   MsgBox Len(vArr)
'   Erase bBBuf
   
   If Err.Number <> 0 Then
      Err.Clear
      On Error GoTo 0
      GetLogicalPages = False
      Exit Function
   End If
   
   ReDim bBBuf(1 To 4)
   bBBuf(1) = Asc("/")
   bBBuf(2) = Asc("P")
   bBBuf(3) = Asc(" ")
   bBBuf(4) = Asc("(")
   i = 1
   Do
      ReDim Preserve FindResult(1 To i)
      If i = 1 Then
         FindResult(i) = FindStringInFile(sFile, bBBuf)
      Else
         FindResult(i) = FindStringInFile(sFile, bBBuf, FindResult(i - 1).lResultPosition + 5)
      End If
      i = i + 1
   Loop Until FindResult(i - 1).bNoErrors = False
   If i > 2 Then
      ReDim Preserve FindResult(1 To i - 2)
   Else ' we DONT have searched string - exiting
      On Error GoTo 0
      GetLogicalPages = False
      Exit Function
   End If
   bBBuf(1) = Asc("/")
   bBBuf(2) = Asc("S")
   bBBuf(3) = Asc("t")
   bBBuf(4) = Asc(" ")
   sTmp = vbNullString
   ReDim sPages(1 To 1)
   ReDim FindResult2(1 To UBound(FindResult))
   For i = 1 To UBound(FindResult)
      sTmp = Replace$(FindResult(i).sResultString, "/P (", vbNullString)
      sTmp = Left$(sTmp, Len(sTmp) - 1)
      FindResult2(i) = FindStringInFile(sFile, bBBuf, FindResult(i).lResultPosition)
      sTmp = sTmp & Replace$(FindResult2(i).sResultString, "/St ", vbNullString)
      sTmp = Replace$(sTmp, ":", vbTab)
      aSplitted = Split(sTmp, vbTab)
      
      'here we must extract sheet number from side number
      lSheetNumber = (aSplitted(1) - 1) \ 2 + 1
      ReDim Preserve aSplitted(0 To 2)
      If aSplitted(1) Mod 2 = 0 Then
         aSplitted(2) = "BACK"
      Else
         aSplitted(2) = "FACE"
      End If
      aSplitted(1) = lSheetNumber
      
      'now we must check separation name - if it is no CMYK then it must be P1
      If UCase$(aSplitted(0)) <> "CYAN" And UCase$(aSplitted(0)) <> "MAGENTA" And _
            UCase$(aSplitted(0)) <> "YELLOW" And UCase$(aSplitted(0)) <> "BLACK" Then
               aSplitted(0) = "P1"
      End If
      sTmp = Join(aSplitted, vbTab)
      sPages(UBound(sPages)) = sTmp
      ReDim Preserve sPages(1 To UBound(sPages) + 1)
      DoEvents
      If bInProcessing = False Then Me.Timer1.Enabled = False: GetLogicalPages = False: Exit Function
   Next i
   
   'sTmp = vbNullString
   'ReDim sPages(1 To 1)
   'For i = LBound(vArr) To UBound(vArr)
   '   If (i < UBound(vArr) - 3) And (vArr(i) Like "/P (*") Then
   '      sTmp = Replace$(vArr(i), "/P (", vbNullString)
   '      sTmp = Left$(sTmp, Len(sTmp) - 1)
   '      i = i + 2
   '      sTmp = sTmp & Replace$(vArr(i), "/St ", vbNullString)
   '      sTmp = Replace$(sTmp, ":", vbTab)
   '      sPages(UBound(sPages)) = sTmp
   '      ReDim Preserve sPages(1 To UBound(sPages) + 1)
   '   End If
   '   If bInProcessing = False Then Me.Timer1.Enabled = False: GetLogicalPages = False: Exit Function
   'Next i
   Erase vArr
   ReDim Preserve sPages(1 To UBound(sPages) - 1)

   If Err.Number = 0 Then GetLogicalPages = True Else Err.Clear: GetLogicalPages = False
   On Error GoTo 0
End Function

Private Function BitsSum(ByVal bByte As Byte) As Byte
   Dim bSum As Byte
   bSum = 0
   bSum = bSum + (bByte And &H1)
   bSum = bSum + (bByte And &H2) \ &H2
   bSum = bSum + (bByte And &H4) \ &H4
   bSum = bSum + (bByte And &H8) \ &H8
   bSum = bSum + (bByte And &H10) \ &H10
   bSum = bSum + (bByte And &H20) \ &H20
   bSum = bSum + (bByte And &H40) \ &H40
   bSum = bSum + (bByte And &H80) \ &H80
   BitsSum = bSum
End Function


Private Function FindStringInFile(ByVal sFileName As String, ByRef bByteArray() As Byte, _
            Optional ByVal lStartPosition As Long = 1) As tFindResult
            
   Const BLOCK_SIZE As Long = 1048576
   Dim FSO As Object, iFN As Integer, lFLen As Long, lNumBlocks As Long, lLastBlockSize As Long
   Dim i As Long, j As Long, k As Long, bReadBuff(1 To BLOCK_SIZE) As Byte, lFindLen As Long
   Dim lLowBoundOfByteArray As Long, lUpBoundOfByteArray As Long, boCompRes As Boolean
   Dim lPosOfBegin As Long, lPosOfEnd As Long, bCurrByte As Byte, sResString As String
   
   If Not IsArray(bByteArray) Then GoTo ERR_HANDLER ' nothing to search!
   lLowBoundOfByteArray = LBound(bByteArray)
   lUpBoundOfByteArray = UBound(bByteArray)
   lFindLen = lUpBoundOfByteArray - lLowBoundOfByteArray + 1
   If lFindLen = 0 Then GoTo ERR_HANDLER ' nothing to search!
   
   On Error Resume Next
   Set FSO = CreateObject("Scripting.FileSystemObject")
   If Err.Number <> 0 Then GoTo ERR_HANDLER ' error accessing ActiveX object!
   If Not FSO.FileExists(sFileName) Then GoTo ERR_HANDLER ' no input file!
   Set FSO = Nothing
   On Error GoTo 0
   
   lFLen = FileLen(sFileName)
   If lStartPosition >= lFLen Then GoTo ERR_HANDLER
   
   lNumBlocks = (lFLen - lStartPosition + 1) \ BLOCK_SIZE ' we need do +1, because first byte number in file is 1, not 0
   lLastBlockSize = (lFLen - lStartPosition + 1) Mod BLOCK_SIZE
   
   iFN = FreeFile()
   Open sFileName For Binary Access Read As iFN
      
      ' lLastBlockSize - we need to expand real file length to read all data
      ' without any additional steps. In fact, exceeded bytes beyond file length
      ' just be filled with ZERO
      For i = lStartPosition To lFLen + (BLOCK_SIZE - lLastBlockSize) Step BLOCK_SIZE
         Erase bReadBuff
         Get #iFN, i, bReadBuff
         For j = 1 To BLOCK_SIZE - lFindLen + 1
            boCompRes = True
            For k = lLowBoundOfByteArray To lUpBoundOfByteArray
               If bReadBuff(j + k - lLowBoundOfByteArray) <> bByteArray(k) Then
                  boCompRes = False
                  Exit For
               End If
'               If k = lUpBoundOfByteArray And boCompRes = True Then
'               End If
            Next k
            If boCompRes Then ' Bingo!!! :)
               lPosOfBegin = i + j ' - lLowBoundOfByteArray
               Exit For
            End If
         Next j
         If boCompRes Then Exit For
         ' then we must shift variable i on lFindLen value
         i = i - lFindLen + 1
      Next i
   
   If boCompRes Then
      FindStringInFile.bNoErrors = True
      
      ' here we step backward until found ZERO, LF or CR symbol
      For i = lPosOfBegin To lStartPosition Step -1
         Get #iFN, i, bCurrByte
         If bCurrByte = &H0 Or bCurrByte = &HA Or bCurrByte = &HD Then
            lPosOfBegin = i + 1
            Exit For
         End If
      Next i
      If i = lStartPosition - 1 Then lPosOfBegin = lStartPosition
      
      ' here we step forward until found ZERO, LF or CR symbol
      For i = lPosOfBegin To lFLen
         Get #iFN, i, bCurrByte
         If bCurrByte = &H0 Or bCurrByte = &HA Or bCurrByte = &HD Then
            lPosOfEnd = i - 1
            Exit For
         End If
      Next i
      If i = lFLen + 1 Then lPosOfEnd = lFLen
      
      sResString = vbNullString
      For i = lPosOfBegin To lPosOfEnd
         Get #iFN, i, bCurrByte
         sResString = sResString & Chr$(bCurrByte)
      Next i
      
      FindStringInFile.lResultPosition = lPosOfBegin
      FindStringInFile.sResultString = sResString
   Else
      FindStringInFile.bNoErrors = False
      FindStringInFile.lResultPosition = 0
      FindStringInFile.sResultString = vbNullString
   End If
   Close #iFN
   Exit Function

ERR_HANDLER:
      Err.Clear
      On Error GoTo 0
      FindStringInFile.bNoErrors = False
      FindStringInFile.lResultPosition = 0
      FindStringInFile.sResultString = vbNullString
End Function

Private Sub txtInkCoeff_Change(Index As Integer)
   If Me.txtInkCoeff(Index) = "." Then
      Me.txtInkCoeff(Index) = "0."
      Me.txtInkCoeff(Index).SelStart = 2
   End If
   If Index < 5 Then
      sMass_Coeff_LIST(Index) = Me.txtInkCoeff(Index).Text
   Else
      sMass_Coeff_ROLL(Index \ 5)(Index Mod 5) = Me.txtInkCoeff(Index).Text
   End If
End Sub

Private Sub txtInkCoeff_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = &H2C Then KeyAscii = &H2E
   If KeyAscii = &H2E And InStr(Me.txtInkCoeff(Index), ".") > 0 Then KeyAscii = &H0
   If (KeyAscii < &H30 Or KeyAscii > &H39) And _
      KeyAscii <> vbKeyBack And KeyAscii <> &H2E Then
         KeyAscii = &H0
   End If
End Sub


