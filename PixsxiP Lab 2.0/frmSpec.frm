VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSpec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pixel Shader Specifics"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView TV 
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   6615
      _Version        =   327682
      Indentation     =   397
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

 'Constants
 TV.Nodes.Add , , "Constants", "Constants"
 TV.Nodes.Add "Constants", tvwChild, "C0", "Constant Registers"
 TV.Nodes.Add "C0", tvwChild, "CR1", "c0"
 TV.Nodes.Add "C0", tvwChild, "CR2", "c1"
 TV.Nodes.Add "C0", tvwChild, "CR3", "c2"
 TV.Nodes.Add "C0", tvwChild, "CR4", "c3"
 TV.Nodes.Add "C0", tvwChild, "CR5", "c4"
 TV.Nodes.Add "C0", tvwChild, "CR6", "c5"
 TV.Nodes.Add "C0", tvwChild, "CR7", "c6"
 TV.Nodes.Add "C0", tvwChild, "CR8", "c7"
 TV.Nodes.Add "Constants", tvwChild, "C1", "Texture Registers"
 TV.Nodes.Add "C1", tvwChild, "TREX", "Note: May go up to t7"
 TV.Nodes.Add "C1", tvwChild, "TR1", "t0"
 TV.Nodes.Add "C1", tvwChild, "TR2", "t1"
 TV.Nodes.Add "Constants", tvwChild, "C2", "Temporary Registers"
 TV.Nodes.Add "C2", tvwChild, "TMPREX", "Note: May go up to r8"
 TV.Nodes.Add "C2", tvwChild, "TMPR1", "r0"
 TV.Nodes.Add "C2", tvwChild, "TMPR2", "r1"
 TV.Nodes.Add "Constants", tvwChild, "C3", "Color Registers"
 TV.Nodes.Add "C3", tvwChild, "CLRR1", "v0"
 TV.Nodes.Add "C3", tvwChild, "CLRR2", "v1"
 DoEvents
 
 'Texture Address Instructions
 TV.Nodes.Add , , "Texture_Address", "Texture Address Instructions"
 TV.Nodes.Add "Texture_Address", tvwChild, "TA1_1", "PS Version 1.1"
 TV.Nodes.Add "TA1_1", tvwChild, "TA111EX1", "d = Destination"
 TV.Nodes.Add "TA1_1", tvwChild, "TA111EX2", "s = Source"
 TV.Nodes.Add "TA1_1", tvwChild, "TA111", "tex (d)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA112", "texbem (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA113", "texbeml (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA114", "texcoord (d)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA115", "texkill (s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA116", "texm3x2pad (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA117", "texm3x2tex (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA118", "texm3x3pad (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA119", "texm3x3spec (d, s1, s2)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA1110", "texm3x3tex (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA1111", "texm3x3vspec (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA1112", "texreg2ar (d, s)"
 TV.Nodes.Add "TA1_1", tvwChild, "TA1113", "texreg2gb (d, s)"
 
 TV.Nodes.Add "Texture_Address", tvwChild, "TA1_2", "PS Version 1.2"
 TV.Nodes.Add "TA1_2", tvwChild, "TA121EX1", "d = Destination"
 TV.Nodes.Add "TA1_2", tvwChild, "TA121EX2", "s = Source"
 TV.Nodes.Add "TA1_2", tvwChild, "TA121", "tex (d)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA122", "texbem (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA123", "texbeml (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA124", "texcoord (d)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA125", "texdp3 (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA126", "texdp3tex (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA127", "texkill (s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA128", "texm3x2pad (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA129", "texm3x2tex (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1210", "texm3x3 (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1211", "texm3x3pad (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1212", "texm3x3spec (d, s1, s2)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1213", "texm3x3tex (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1214", "texm3x3vspec (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1215", "texreg2ar (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1216", "texreg2gb (d, s)"
 TV.Nodes.Add "TA1_2", tvwChild, "TA1217", "texreg2rgb (d, s)"
 
 TV.Nodes.Add "Texture_Address", tvwChild, "TA1_3", "PS Version 1.3"
 TV.Nodes.Add "TA1_3", tvwChild, "TA131EX1", "d = Destination"
 TV.Nodes.Add "TA1_3", tvwChild, "TA131EX2", "s = Source"
 TV.Nodes.Add "TA1_3", tvwChild, "TA131", "tex (d)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA132", "texbem (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA133", "texbeml (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA134", "texcoord (d)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA135", "texdp3 (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA136", "texdp3tex (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA137", "texkill (s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA138", "texm3x2depth (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA139", "texm3x2pad (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1310", "texm3x2tex (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1311", "texm3x3 (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1312", "texm3x3pad (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1313", "texm3x3spec (d, s1, s2)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1314", "texm3x3tex (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1315", "texm3x3vspec (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1316", "texreg2ar (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1317", "texreg2gb (d, s)"
 TV.Nodes.Add "TA1_3", tvwChild, "TA1318", "texreg2rgb (d, s)"
 
 TV.Nodes.Add "Texture_Address", tvwChild, "TA1_4", "PS Version 1.4"
 TV.Nodes.Add "TA1_4", tvwChild, "TA141EX1", "d = Destination"
 TV.Nodes.Add "TA1_4", tvwChild, "TA141EX2", "s = Source"
 TV.Nodes.Add "TA1_4", tvwChild, "TA141", "texcrd (d, s)"
 TV.Nodes.Add "TA1_4", tvwChild, "TA142", "texdepth (d)"
 TV.Nodes.Add "TA1_4", tvwChild, "TA143", "texkill (s)"
 TV.Nodes.Add "TA1_4", tvwChild, "TA144", "texld (d, s)"
 DoEvents
 
 'Arithmetic Instructions
 TV.Nodes.Add , , "Arithmetic", "Arithmetic Instructions"
 TV.Nodes.Add "Arithmetic", tvwChild, "A1_1", "PS Version 1.1 - 1.3"
  '
  
 TV.Nodes.Add "Arithmetic", tvwChild, "A1_4", "PS Version 1.4"
  '
 DoEvents
 
 SetParent frmSpec.hWnd, frmMain.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.chkSpec.Value = 0
End Sub
