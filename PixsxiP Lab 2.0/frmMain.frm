VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PixsxiP Lab 2.0"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   771
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView TV 
      Height          =   1275
      Left            =   0
      TabIndex        =   35
      Top             =   6915
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   2249
      _Version        =   327682
      Indentation     =   397
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Dummy 
      Caption         =   "Check1"
      Height          =   495
      Left            =   5160
      TabIndex        =   34
      Top             =   12000
      Width           =   1215
   End
   Begin VB.Frame fmControlsCont 
      BackColor       =   &H00E0E0E0&
      Height          =   6885
      Left            =   6750
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox picTexCont 
         BackColor       =   &H00E0E0E0&
         Height          =   1380
         Left            =   80
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   307
         TabIndex        =   36
         Top             =   2470
         Width           =   4670
         Begin VB.PictureBox picTex 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   240
            ScaleHeight     =   3855
            ScaleWidth      =   3855
            TabIndex        =   38
            Top             =   0
            Width           =   3855
            Begin VB.TextBox txtTex 
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   50
               Top             =   360
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Height          =   270
               Index           =   0
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox txtTex 
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   48
               Top             =   960
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   270
               Index           =   1
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   960
               Width           =   255
            End
            Begin VB.TextBox txtTex 
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   46
               Top             =   1560
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   270
               Index           =   2
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   1560
               Width           =   255
            End
            Begin VB.TextBox txtTex 
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   44
               Top             =   2160
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   270
               Index           =   3
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   2160
               Width           =   255
            End
            Begin VB.TextBox txtTex 
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   42
               Top             =   2760
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   270
               Index           =   4
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   2760
               Width           =   255
            End
            Begin VB.TextBox txtTex 
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   40
               Top             =   3360
               Width           =   3375
            End
            Begin VB.CommandButton cmdOpenText 
               BackColor       =   &H00E0E0E0&
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   270
               Index           =   5
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   3360
               Width           =   255
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   56
               Top             =   120
               Width           =   780
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   120
               TabIndex        =   55
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   120
               TabIndex        =   54
               Top             =   1320
               Width           =   780
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   120
               TabIndex        =   53
               Top             =   1920
               Width           =   780
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   120
               TabIndex        =   52
               Top             =   2520
               Width           =   780
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Texture 6"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   120
               TabIndex        =   51
               Top             =   3120
               Width           =   780
            End
         End
         Begin VB.VScrollBar TexScroll 
            Height          =   1335
            LargeChange     =   75
            Left            =   4410
            Max             =   169
            SmallChange     =   50
            TabIndex        =   37
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.CheckBox chkSpec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "View Specifics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1380
         Width           =   975
      End
      Begin RichTextLib.RichTextBox txtShader 
         Height          =   2260
         Left            =   120
         TabIndex        =   31
         Top             =   4230
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   3995
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":0CCE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCompile 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
         Width           =   975
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Load Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Text            =   "0.15"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Text            =   "0.75"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Text            =   "0.25"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc0 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Text            =   "0.15"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Text            =   "1"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Text            =   "0.5"
         Top             =   1920
         Width           =   500
      End
      Begin VB.TextBox txtc1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   8
         Text            =   "0"
         Top             =   1920
         Width           =   500
      End
      Begin VB.CommandButton cmdSamBack 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4150
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton cmdSamFor 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4180
         TabIndex        =   58
         Top             =   550
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   0
         Picture         =   "frmMain.frx":0D96
         Top             =   105
         Width           =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   3610
         X2              =   3610
         Y1              =   800
         Y2              =   2400
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   3600
         X2              =   3600
         Y1              =   800
         Y2              =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixel Shader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   3990
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4800
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4800
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4800
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4800
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constant c0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constant c1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   27
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1050
         TabIndex        =   26
         Top             =   1230
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1890
         TabIndex        =   25
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   24
         Top             =   1230
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   23
         Top             =   1950
         Width           =   105
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1050
         TabIndex        =   22
         Top             =   1950
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1890
         TabIndex        =   21
         Top             =   1950
         Width           =   105
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   20
         Top             =   1950
         Width           =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4800
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4800
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblSample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample - 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3135
         TabIndex        =   19
         Top             =   3990
         Width           =   855
      End
   End
   Begin VB.Frame fmCanvasCont 
      BackColor       =   &H00E0E0E0&
      Height          =   5610
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      Begin VB.PictureBox Canvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   5300
         Left            =   100
         ScaleHeight     =   349
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   429
         TabIndex        =   1
         Top             =   200
         Width           =   6500
         Begin VB.CommandButton cmdAbout 
            BackColor       =   &H00E0E0E0&
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5475
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   5040
            Width           =   975
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   5040
            Width           =   975
         End
      End
   End
   Begin VB.Frame fmErrorCont 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   5550
      Width           =   6720
      Begin VB.TextBox txtError 
         BackColor       =   &H00C0C0C0&
         Height          =   900
         Left            =   80
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   6570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DirectX Assembler Result"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   4
         Top             =   120
         Width           =   2130
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSpec_Click()

 Dummy.SetFocus
 If chkSpec.Value Then
  frmMain.Top = 0: frmMain.Height = 8565
 Else
  frmMain.Top = 855: frmMain.Height = 7275
 End If
 
End Sub

Private Sub cmdAbout_Click()
 frmAbout.Show vbModal, frmMain
 DoEvents
 UpdateShader
 Render
End Sub

Private Sub cmdCompile_Click()
 
 Dummy.SetFocus
 If Not HasTexture1 Then MsgBox "Please load a texture...", vbInformation, "PixsxiP Lab Information": Exit Sub
 Screen.MousePointer = 11
 
 If PS_Handle Then
  Call D3DD.DeletePixelShader(PS_Handle)
  PS_Handle = 0
 End If
 
 PixelShader = txtShader.Text
 UpdateShader
 Render
 Screen.MousePointer = 0
 
End Sub

Private Sub cmdLoad_Click()
 Dim tmpString As String
 Dim FSys As New FileSystemObject
 Dim InputStream As TextStream
 
 Dummy.SetFocus
 CD.Filter = "Vertex Shader File (*.txt)|*.txt"
 CD.DialogTitle = "Open Pixel Shader"
 CD.InitDir = App.Path
 CD.FileName = ""
 CD.ShowOpen
 
 If CD.FileName <> "" Then
  txtShader.Text = ""
  Set InputStream = FSys.OpenTextFile(CD.FileName)
  
  Do Until InputStream.AtEndOfStream = True
   tmpString = InputStream.ReadLine
   txtShader.Text = txtShader.Text & tmpString & vbNewLine
  Loop
 End If
 
 Set InputStream = Nothing
 Set FSys = Nothing
End Sub

Private Sub cmdOpenText_Click(Index As Integer)
 
 Dummy.SetFocus
 CD.Filter = "Bitmap Image (*.bmp)|*.bmp|TGA Image (*.tga)|*.tga|JPG Image (*.jpg)|*.jpg"
 CD.DialogTitle = "Open Texture" & Index
 CD.InitDir = App.Path
 CD.FileName = ""
 CD.ShowOpen
 
 If CD.FileName <> "" Then
  txtTex(Index).Text = CD.FileTitle
  Select Case Index
   Case 0: HasTexture1 = True
  End Select
  
  Set Mesh.Texture(Index) = Nothing
  Set Mesh.Texture(Index) = D3DX.CreateTextureFromFile(D3DD, CD.FileName)
 End If
 
End Sub

Private Sub cmdRefresh_Click()
 Dummy.SetFocus
 Screen.MousePointer = 11
 Render
 Screen.MousePointer = 0
End Sub

Private Sub cmdSamBack_Click()
 Dummy.SetFocus
 ToggleSampleCode Backward
End Sub

Private Sub cmdSamFor_Click()
 Dummy.SetFocus
 ToggleSampleCode Forward
End Sub

Private Sub cmdSave_Click()
 Dim tmpString As String
 Dim FSys As New FileSystemObject
 Dim OutputStream As TextStream
 
 Dummy.SetFocus
 CD.Filter = "Vertex Shader File (*.txt)|*.txt"
 CD.DialogTitle = "Open Pixel Shader"
 CD.FileName = ""
 CD.flags = &HF
 CD.ShowSave
 
 If CD.FileName <> "" Then
  Set OutputStream = FSys.CreateTextFile(CD.FileName)
  OutputStream.Write txtShader.Text
 End If
 
 Set OutputStream = Nothing
 Set FSys = Nothing
End Sub

Private Sub Form_Load()
 Dim i As Long
 
 SetupTreeView
 MakeSamples
 PixelShader = txtShader.Text
 txtShader.Text = Samples(0)
 Initialize
 
 For i = 0 To 5
  txtTex(i).Text = "Video Card Supports " & MaxVideoCardTextures & " Simultaneous Textures"
  txtTex(i).Enabled = False
  cmdOpenText(i).Enabled = False
  Set Mesh.Texture(i) = Nothing
 Next
 
 For i = 0 To MaxVideoCardTextures - 1
  txtTex(i).Text = ""
  txtTex(i).Enabled = True
  cmdOpenText(i).Enabled = True
 Next
 
End Sub

Private Sub Form_Resize()
 Dummy.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Unload frmAbout
 Cleanup_DX8
End Sub

Private Sub TexScroll_Change()
 picTex.Top = -TexScroll.Value
End Sub

Private Sub TexScroll_Scroll()
 picTex.Top = -TexScroll.Value
End Sub

Private Sub txtc0_LostFocus(Index As Integer)
 Dim tmpString As String
 
 If Left$(txtc0(Index).Text, 1) = "." Then
  tmpString = txtc0(Index).Text
  txtc0(Index).Text = CSng("0" & tmpString)
 ElseIf Left$(txtc0(Index).Text, 2) = "-." Then
  tmpString = Mid$(txtc0(Index).Text, 2, 50)
  txtc0(Index).Text = CSng("-0" & Trim$(tmpString))
 ElseIf CSng(txtc0(Index).Text) > 1 Then
  txtc0(Index).Text = 1
 ElseIf CSng(txtc0(Index).Text) < -1 Then
  txtc0(Index).Text = -1
 End If
 
End Sub

Private Sub txtc1_LostFocus(Index As Integer)
 Dim tmpString As String
 
 If Left$(txtc1(Index).Text, 1) = "." Then
  tmpString = txtc1(Index).Text
  txtc1(Index).Text = CSng("0" & tmpString)
 ElseIf Left$(txtc1(Index).Text, 2) = "-." Then
  tmpString = Mid$(txtc1(Index).Text, 2, 50)
  txtc1(Index).Text = CSng("-0" & Trim$(tmpString))
 ElseIf CSng(txtc1(Index).Text) > 1 Then
  txtc1(Index).Text = 1
 ElseIf CSng(txtc1(Index).Text) < -1 Then
  txtc1(Index).Text = -1
 End If
 
End Sub
