VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   4695
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
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
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utility"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1500
         TabIndex        =   5
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixel Shader"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   855
         TabIndex        =   4
         Top             =   600
         Width           =   3060
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "©2002 Ariel Productions"
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
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   1965
      End
   End
   Begin VB.Label Label1 
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
      Left            =   4170
      TabIndex        =   3
      Top             =   405
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
 Unload frmAbout
End Sub

