VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "WaynhSoft - Restaurant & Bar"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13515
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data4 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   7920
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2880
      Top             =   720
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "BAR"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   6615
      Index           =   1
      Left            =   6840
      TabIndex        =   35
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":0CCA
         Height          =   855
         Index           =   1
         Left            =   240
         MouseIcon       =   "Form1.frx":264C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2956
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 2"
         DownPicture     =   "Form1.frx":42D8
         Height          =   855
         Index           =   2
         Left            =   1440
         MouseIcon       =   "Form1.frx":5C5A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":5F64
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":78E6
         Height          =   855
         Index           =   3
         Left            =   2640
         MouseIcon       =   "Form1.frx":9268
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9572
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":AEF4
         Height          =   855
         Index           =   4
         Left            =   3840
         MouseIcon       =   "Form1.frx":C876
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":CB80
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":E502
         Height          =   855
         Index           =   5
         Left            =   5040
         MouseIcon       =   "Form1.frx":FE84
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1018E
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":11B10
         Height          =   855
         Index           =   6
         Left            =   240
         MouseIcon       =   "Form1.frx":13492
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1379C
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":1511E
         Height          =   855
         Index           =   7
         Left            =   1440
         MouseIcon       =   "Form1.frx":16AA0
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":16DAA
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":1872C
         Height          =   855
         Index           =   8
         Left            =   2640
         MouseIcon       =   "Form1.frx":1A0AE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1A3B8
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":1BD3A
         Height          =   855
         Index           =   9
         Left            =   3840
         MouseIcon       =   "Form1.frx":1D6BC
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1D9C6
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":1F348
         Height          =   855
         Index           =   10
         Left            =   5040
         MouseIcon       =   "Form1.frx":20CCA
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":20FD4
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":22956
         Height          =   855
         Index           =   11
         Left            =   240
         MouseIcon       =   "Form1.frx":242D8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":245E2
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":25F64
         Height          =   855
         Index           =   12
         Left            =   1440
         MouseIcon       =   "Form1.frx":278E6
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":27BF0
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":29572
         Height          =   855
         Index           =   13
         Left            =   2640
         MouseIcon       =   "Form1.frx":2AEF4
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2B1FE
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":2CB80
         Height          =   855
         Index           =   14
         Left            =   3840
         MouseIcon       =   "Form1.frx":2E502
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2E80C
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":3018E
         Height          =   855
         Index           =   15
         Left            =   5040
         MouseIcon       =   "Form1.frx":31B10
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":31E1A
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":3379C
         Height          =   855
         Index           =   16
         Left            =   240
         MouseIcon       =   "Form1.frx":3511E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":35428
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":36DAA
         Height          =   855
         Index           =   17
         Left            =   1440
         MouseIcon       =   "Form1.frx":3872C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":38A36
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":3A3B8
         Height          =   855
         Index           =   18
         Left            =   2640
         MouseIcon       =   "Form1.frx":3BD3A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":3C044
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":3D9C6
         Height          =   855
         Index           =   19
         Left            =   3840
         MouseIcon       =   "Form1.frx":3F348
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":3F652
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":40FD4
         Height          =   855
         Index           =   20
         Left            =   5040
         MouseIcon       =   "Form1.frx":42956
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":42C60
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":445E2
         Height          =   855
         Index           =   21
         Left            =   240
         MouseIcon       =   "Form1.frx":45F64
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4626E
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":47BF0
         Height          =   855
         Index           =   22
         Left            =   1440
         MouseIcon       =   "Form1.frx":49572
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4987C
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":4B1FE
         Height          =   855
         Index           =   23
         Left            =   2640
         MouseIcon       =   "Form1.frx":4CB80
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":4CE8A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":4E80C
         Height          =   855
         Index           =   24
         Left            =   3840
         MouseIcon       =   "Form1.frx":5018E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":50498
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":51E1A
         Height          =   855
         Index           =   25
         Left            =   5040
         MouseIcon       =   "Form1.frx":5379C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":53AA6
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":55428
         Height          =   855
         Index           =   26
         Left            =   240
         MouseIcon       =   "Form1.frx":56DAA
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":570B4
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":58A36
         Height          =   855
         Index           =   27
         Left            =   1440
         MouseIcon       =   "Form1.frx":5A3B8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":5A6C2
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":5C044
         Height          =   855
         Index           =   28
         Left            =   2640
         MouseIcon       =   "Form1.frx":5D9C6
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":5DCD0
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":5F652
         Height          =   855
         Index           =   29
         Left            =   3840
         MouseIcon       =   "Form1.frx":60FD4
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":612DE
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":62C60
         Height          =   855
         Index           =   30
         Left            =   5040
         MouseIcon       =   "Form1.frx":645E2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":648EC
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5520
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList3 
         Index           =   1
         Left            =   6120
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6626E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":67C00
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":69592
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   6120
         Y1              =   4440
         Y2              =   4440
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   13455
      TabIndex        =   33
      Top             =   7635
      Width           =   13515
      Begin VB.Timer Timer3 
         Interval        =   100
         Left            =   3960
         Top             =   360
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   1560
         MouseIcon       =   "Form1.frx":6AF24
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":6B22E
         ToolTipText     =   "MiSIiii nUmPAnG LEwaT..."
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name : Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   600
         MouseIcon       =   "Form1.frx":6CEF8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":6D202
         ToolTipText     =   "Log Off"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   120
         MouseIcon       =   "Form1.frx":6E284
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":6E58E
         ToolTipText     =   "Ganti Password"
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "RESTAURANT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":6F610
         Height          =   855
         Index           =   29
         Left            =   5040
         MouseIcon       =   "Form1.frx":70F92
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7129C
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":72C1E
         Height          =   855
         Index           =   28
         Left            =   3840
         MouseIcon       =   "Form1.frx":745A0
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":748AA
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":7622C
         Height          =   855
         Index           =   27
         Left            =   2640
         MouseIcon       =   "Form1.frx":77BAE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":77EB8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":7983A
         Height          =   855
         Index           =   26
         Left            =   1440
         MouseIcon       =   "Form1.frx":7B1BC
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7B4C6
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":7CE48
         Height          =   855
         Index           =   25
         Left            =   240
         MouseIcon       =   "Form1.frx":7E7CA
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":7EAD4
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":80456
         Height          =   855
         Index           =   24
         Left            =   5040
         MouseIcon       =   "Form1.frx":81DD8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":820E2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":83A64
         Height          =   855
         Index           =   23
         Left            =   3840
         MouseIcon       =   "Form1.frx":853E6
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":856F0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":87072
         Height          =   855
         Index           =   22
         Left            =   2640
         MouseIcon       =   "Form1.frx":889F4
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":88CFE
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":8A680
         Height          =   855
         Index           =   21
         Left            =   1440
         MouseIcon       =   "Form1.frx":8C002
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":8C30C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":8DC8E
         Height          =   855
         Index           =   20
         Left            =   240
         MouseIcon       =   "Form1.frx":8F610
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":8F91A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":9129C
         Height          =   855
         Index           =   19
         Left            =   5040
         MouseIcon       =   "Form1.frx":92C1E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":92F28
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":948AA
         Height          =   855
         Index           =   18
         Left            =   3840
         MouseIcon       =   "Form1.frx":9622C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":96536
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":97EB8
         Height          =   855
         Index           =   17
         Left            =   2640
         MouseIcon       =   "Form1.frx":9983A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":99B44
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":9B4C6
         Height          =   855
         Index           =   16
         Left            =   1440
         MouseIcon       =   "Form1.frx":9CE48
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9D152
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":9EAD4
         Height          =   855
         Index           =   15
         Left            =   240
         MouseIcon       =   "Form1.frx":A0456
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A0760
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":A20E2
         Height          =   855
         Index           =   14
         Left            =   5040
         MouseIcon       =   "Form1.frx":A3A64
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A3D6E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":A56F0
         Height          =   855
         Index           =   13
         Left            =   3840
         MouseIcon       =   "Form1.frx":A7072
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A737C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":A8CFE
         Height          =   855
         Index           =   12
         Left            =   2640
         MouseIcon       =   "Form1.frx":AA680
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":AA98A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":AC30C
         Height          =   855
         Index           =   11
         Left            =   1440
         MouseIcon       =   "Form1.frx":ADC8E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":ADF98
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":AF91A
         Height          =   855
         Index           =   10
         Left            =   240
         MouseIcon       =   "Form1.frx":B129C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":B15A6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":B2F28
         Height          =   855
         Index           =   9
         Left            =   5040
         MouseIcon       =   "Form1.frx":B48AA
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":B4BB4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":B6536
         Height          =   855
         Index           =   8
         Left            =   3840
         MouseIcon       =   "Form1.frx":B7EB8
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":B81C2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":B9B44
         Height          =   855
         Index           =   7
         Left            =   2640
         MouseIcon       =   "Form1.frx":BB4C6
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":BB7D0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":BD152
         Height          =   855
         Index           =   6
         Left            =   1440
         MouseIcon       =   "Form1.frx":BEAD4
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":BEDDE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":C0760
         Height          =   855
         Index           =   5
         Left            =   240
         MouseIcon       =   "Form1.frx":C20E2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":C23EC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA"
         DownPicture     =   "Form1.frx":C3D6E
         Height          =   855
         Index           =   4
         Left            =   5040
         MouseIcon       =   "Form1.frx":C56F0
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":C59FA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA"
         DownPicture     =   "Form1.frx":C737C
         Height          =   855
         Index           =   3
         Left            =   3840
         MouseIcon       =   "Form1.frx":C8CFE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":C9008
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":CA98A
         Height          =   855
         Index           =   2
         Left            =   2640
         MouseIcon       =   "Form1.frx":CC30C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":CC616
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 2"
         DownPicture     =   "Form1.frx":CDF98
         Height          =   855
         Index           =   1
         Left            =   1440
         MouseIcon       =   "Form1.frx":CF91A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":CFC24
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList3 
         Index           =   0
         Left            =   6120
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":D15A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":D2F38
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":D48CA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton MEJA 
         BackColor       =   &H0000FFFF&
         Caption         =   "MEJA 1"
         DownPicture     =   "Form1.frx":D625C
         Height          =   855
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form1.frx":D7BDE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":D7EE8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   240
         X2              =   6120
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   240
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   8625
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   18177
            Picture         =   "Form1.frx":D986A
            Text            =   "WaynhSoft "
            TextSave        =   "WaynhSoft "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   3942
            Picture         =   "Form1.frx":D9FE4
            Text            =   "www.WaynhSoft.co.cc   "
            TextSave        =   "www.WaynhSoft.co.cc   "
            Object.ToolTipText     =   "Website"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   4154
            Picture         =   "Form1.frx":DB076
            Text            =   "021-2626 4 888 (Wahyu)   "
            TextSave        =   "021-2626 4 888 (Wahyu)   "
            Object.ToolTipText     =   "Telepon"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   4339
            Picture         =   "Form1.frx":DB7F0
            Text            =   "Waynh@WaynhSoft.co.cc  "
            TextSave        =   "Waynh@WaynhSoft.co.cc  "
            Object.ToolTipText     =   "Email"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   5054
            Picture         =   "Form1.frx":DC882
            Text            =   "Wahyu_NHidayat@yahoo.com  "
            TextSave        =   "Wahyu_NHidayat@yahoo.com  "
            Object.ToolTipText     =   "Email/Yahoo Messenger"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Picture         =   "Form1.frx":DCB9C
            TextSave        =   "4:16 AM"
            Object.ToolTipText     =   "Jam"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Picture         =   "Form1.frx":DD5AE
            TextSave        =   "4/15/2009"
            Object.ToolTipText     =   "Tanggal"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   1217
      ButtonWidth     =   1879
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PRODUK"
            Object.ToolTipText     =   "Database Produk"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LAPORAN"
            Object.ToolTipText     =   "Laporan Penjualan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "KALKULATOR"
            Object.ToolTipText     =   "Kalkulator"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "MANUAL"
            Object.ToolTipText     =   "User Manual"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Object.ToolTipText     =   "Keluar Dari Program"
            ImageIndex      =   5
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":DDFC0
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":DE2DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":DFC6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E15FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E2F90
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E4922
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E55FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E6F8E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E8920
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E909A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E9814
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E9F8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":EA708
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":EAE82
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":EB5FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":EBD76
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":EC4F0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu produk_mnu 
      Caption         =   "&Produk"
   End
   Begin VB.Menu edit_mnu 
      Caption         =   "&Edit Data"
      Visible         =   0   'False
   End
   Begin VB.Menu ctkUlang_mnu 
      Caption         =   "&Cetak Ulang"
      Visible         =   0   'False
   End
   Begin VB.Menu Laporan_mnu 
      Caption         =   "&Laporan Penjualan"
   End
   Begin VB.Menu sett_mnu 
      Caption         =   "&Setting %"
   End
   Begin VB.Menu user_mnu 
      Caption         =   "&User Account"
   End
   Begin VB.Menu kalk_mnu 
      Caption         =   "&Kalkulator"
   End
   Begin VB.Menu manual_mnu 
      Caption         =   "User Manual"
      Visible         =   0   'False
   End
   Begin VB.Menu x_mnu 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txt1 As String
Dim con1 As Byte
Dim txt2 As String
Dim con2 As Byte

Private Sub Command1_Click(Index As Integer)
With Trans_frm
    .Caption = "BAR " & Index
    .Label1(0).Caption = "TRANSAKSI BAR " & Index
End With
Me.Visible = False
Trans_frm.Show
End Sub

Private Sub Form_Activate()
button_awal
Call db_kontrol
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Mid(!nama_table, 1, 4) = "MEJA" Then
            MEJA(Val(Mid(!nama_table, 5, 3)) - 1).Picture = ImageList3(0).ListImages(3).Picture
        Else
            Command1(Val(Mid(!nama_table, 4, 3))).Picture = ImageList3(1).ListImages(2).Picture
        End If
        .MoveNext
    Loop
    End If
End With
End Sub

Sub button_awal()
Dim x As Byte
x = 0
Do Until x = 30
    MEJA(x).Caption = "MEJA " & x + 1
    MEJA(x).Picture = ImageList3(0).ListImages(1).Picture
    x = x + 1
Loop
x = 1
Do Until x = 31
    Command1(x).Caption = "BAR " & x
    Command1(x).Picture = ImageList3(1).ListImages(1).Picture
    x = x + 1
Loop
End Sub

Private Sub Form_Load()
con1 = 0
con2 = 0
Image2.Left = -600
Randomize
End Sub

Private Sub Form_Unload(Cancel As Integer)
x = MsgBox("Apakah anda yakin keluar dari Aplikasi...???", vbYesNo, "Exit Program")
If x = vbYes Then
    End
Else
    Cancel = -1
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    x = MsgBox("Apakah anda yakin Log Off", vbYesNo, "LOG OFF")
    If x = vbYes Then
        Me.Hide
        frmLogin.Show
    End If
Case 1
    Me.Enabled = False
    With Password_frm
        .Data1.DatabaseName = Data4.DatabaseName
        .Data1.RecordSource = Data4.RecordSource
        .Label1(13).Caption = Label1
        .Show
    End With
End Select
End Sub

Private Sub Image2_Click()
Dim x As Integer
x = Int(Rnd * 10)
If x = 1 Then
    MsgBox "Ayo kerja yang bener...", vbInformation, "Pesan"
ElseIf x = 2 Then
    MsgBox "SEMANGAAAAAT...", vbInformation, "Pesan"
ElseIf x = 3 Then
    MsgBox "HaLLoo ApA kABaR...", vbInformation, "Pesan"
ElseIf x = 4 Then
    MsgBox "KeJAR sEToRan...", vbInformation, "Pesan"
ElseIf x = 5 Then
    MsgBox "nGApAin SiH CoLEK2 aKYuu...", vbInformation, "Pesan"
ElseIf x = 6 Then
    MsgBox "NgGAK bOleH MaLEZzzz...", vbInformation, "Pesan"
ElseIf x = 7 Then
    MsgBox "KoNSENtrASiii...", vbInformation, "Pesan"
ElseIf x = 8 Then
    MsgBox "MiSii NumPang LeWat...", vbInformation, "Pesan"
ElseIf x = 9 Then
    MsgBox "PuRA-pUra SIbUK...", vbInformation, "Pesan"
Else
    MsgBox "UDaH mAKAn BeLUm...", vbInformation, "Pesan"
End If
End Sub

Private Sub kalk_mnu_Click()
    AppActivate Shell("calc.exe", 1)
End Sub

Private Sub Laporan_mnu_Click()
lap = True
Call passwd
End Sub

Private Sub MEJA_Click(Index As Integer)
With Trans_frm
    .Caption = "MEJA " & Index + 1
    .Label1(0).Caption = "TRANSAKSI MEJA " & Index + 1
End With
Trans_frm.Show
Me.Visible = False
End Sub

Private Sub produk_mnu_Click()
prd = True
Call passwd
End Sub

Private Sub sett_mnu_Click()
sett = True
Call passwd
End Sub

Private Sub Timer1_Timer()
Dim car As Single
txt1 = "RESTAURANT"
car = Len(txt1)
If con1 < car Then
    Frame1(0).Caption = Mid(txt1, 1, con1 + 1)
    con1 = con1 + 1
Else
    con1 = 0
End If
End Sub

Private Sub Timer2_Timer()
txt2 = "BAR"
If Frame1(1).Caption = "" Then
    Frame1(1).Caption = txt2
Else
    Frame1(1).Caption = ""
End If
End Sub

Private Sub Timer3_Timer()
If Image2.Left < 13440 Then
    Image2.Left = Image2.Left + 50
Else
    Image2.Left = -600
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
produk_mnu_Click
Case 2
Laporan_mnu_Click
Case 3
kalk_mnu_Click
Case 4
Case 5
    Unload Me
End Select
End Sub

Private Sub user_mnu_Click()
use = True
Call passwd
End Sub

Private Sub x_mnu_Click()
Unload Me
End Sub
