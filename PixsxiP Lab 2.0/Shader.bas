Attribute VB_Name = "Shader"
Option Explicit

Public Sub UpdateShader()
 On Local Error Resume Next
 
 PS_Constants(0, 0) = CSng(frmMain.txtc0(0).Text)
 PS_Constants(1, 0) = CSng(frmMain.txtc0(1).Text)
 PS_Constants(2, 0) = CSng(frmMain.txtc0(2).Text)
 PS_Constants(3, 0) = CSng(frmMain.txtc0(3).Text)
 PS_Constants(0, 1) = CSng(frmMain.txtc1(0).Text)
 PS_Constants(1, 1) = CSng(frmMain.txtc1(1).Text)
 PS_Constants(2, 1) = CSng(frmMain.txtc1(2).Text)
 PS_Constants(3, 1) = CSng(frmMain.txtc1(3).Text)
 CreateShaderFromCode
 
End Sub

Public Sub CreateShaderFromCode()
 On Local Error Resume Next
 Dim ShaderCode As D3DXBuffer
 Dim RetError As String
 Dim Arr() As Long
 Dim Size As Long
 Dim i As Long
  
 Set ShaderCode = D3DX.AssembleShader(PixelShader, 0, Nothing)
 Size = ShaderCode.GetBufferSize() / 4
 ReDim Arr(Size - 1)
 D3DX.BufferGetData ShaderCode, 0, 4, Size, Arr(0)
 PS_Handle = D3DD.CreatePixelShader(Arr(0))
 
 If PS_Handle Then
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "Compiled Successfully"
 Else
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "Error..."
 End If
 If PixelShader = "" Then
  frmMain.txtError.Text = ""
  frmMain.txtError.Text = "No Pixel Shader Defined..."
 End If
 
 Set ShaderCode = Nothing
 DoEvents
 
 D3DD.SetStreamSource 0, DX_VB, Len(Mesh.Vertices(0))
 D3DD.SetVertexShader FVF_ShaderVertex
 D3DD.SetPixelShaderConstant 0, PS_Constants(0, 0), 2
End Sub

'Here down are just some samples.

Public Sub MakeSamples()
 'I set this up for readability
 
 Samples(0) = ";------------------------------------" & vbNewLine & _
              "; Show Texture 0                     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "mov r0, t0                           "
              
 Samples(1) = ";------------------------------------" & vbNewLine & _
              "; Show Texture 1                     " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t1                           "
              
 Samples(2) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Texture 0               " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t0                           " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(3) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Texture 1               " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r0, t1                           " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(4) = ";------------------------------------" & vbNewLine & _
              "; Blend_V Textures 1                 " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t1                           " & vbNewLine & _
              "lrp r0, v1, r1, t0                   "
              
 Samples(5) = ";------------------------------------" & vbNewLine & _
              "; Blend_V Textures 2                 " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t0                           " & vbNewLine & _
              "lrp r0, v1, r1, t1                   "
              
 Samples(6) = ";------------------------------------" & vbNewLine & _
              "; Blend_H Textures 1                 " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t1                           " & vbNewLine & _
              "lrp r0, v0, r1, t0                   "
              
 Samples(7) = ";------------------------------------" & vbNewLine & _
              "; Blend_H Textures 2                 " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t0                           " & vbNewLine & _
              "lrp r0, v0, r1, t1                   "
              
 Samples(8) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Blend_V Textures 1      " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t1                           " & vbNewLine & _
              "lrp r0, v1, r1, t0                   " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(9) = ";------------------------------------" & vbNewLine & _
              "; Gray Scale Blend_V Textures 2      " & vbNewLine & _
              ";------------------------------------" & vbNewLine & _
              "ps.1.0                               " & vbNewLine & _
              "                                     " & vbNewLine & _
              "tex t0                               " & vbNewLine & _
              "tex t1                               " & vbNewLine & _
              "mov r1, t0                           " & vbNewLine & _
              "lrp r0, v1, r1, t1                   " & vbNewLine & _
              "dp3 r0, r0, c0                       "
              
 Samples(10) = ";-----------------------------------" & vbNewLine & _
               "; Gray Scale Blend_H Textures 1     " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.0                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mov r1, t1                          " & vbNewLine & _
               "lrp r0, v0, r1, t0                  " & vbNewLine & _
               "dp3 r0, r0, c0                      "
              
 Samples(11) = ";-----------------------------------" & vbNewLine & _
               "; Gray Scale Blend_H Textures 2     " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.0                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mov r1, t0                          " & vbNewLine & _
               "lrp r0, v0, r1, t1                  " & vbNewLine & _
               "dp3 r0, r0, c0                      "
               
 Samples(12) = ";-----------------------------------" & vbNewLine & _
               "; Color Blend Textures 0            " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.0                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "mov r1, t0                          " & vbNewLine & _
               "lrp r0, v1, r1, c0                  "
               
 Samples(13) = ";-----------------------------------" & vbNewLine & _
               "; Color Blend Textures 1            " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.0                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mov r1, t1                          " & vbNewLine & _
               "lrp r0, v1, r1, c0                  "
               
 Samples(14) = ";-----------------------------------" & vbNewLine & _
               "; Alpha-Red UV                      " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "texreg2ar t1, t0                    " & vbNewLine & _
               "mov r0, t1                          "
               
               
 Samples(15) = ";-----------------------------------" & vbNewLine & _
               "; Green-Blue UV                     " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "texreg2gb t1, t0                    " & vbNewLine & _
               "mov r0, t1                          "
               
 Samples(16) = ";-----------------------------------" & vbNewLine & _
               "; Red-Green-Blue UV                 " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.2                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "texreg2rgb t1, t0                   " & vbNewLine & _
               "mov r0, t1                          "
               
 Samples(17) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Textures 0           " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "add r0, v1, t0                      "
               
 Samples(18) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Textures 1           " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add r0, v1, t1                      "
               
 Samples(19) = ";-----------------------------------" & vbNewLine & _
               "; Color Glow Mapping Textures 0     " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "add r0, v0, t0                      "
               
 Samples(20) = ";-----------------------------------" & vbNewLine & _
               "; Color Glow Mapping Textures 1     " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add r0, v0, t1                      "
               
 Samples(21) = ";-----------------------------------" & vbNewLine & _
               "; Color Glow Mapping Blend          " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add r0, t0, t1                      "
               
 Samples(22) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Blend X2             " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add_x2 r0, t0, t1                   "
               
 Samples(23) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Blend X4             " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add_x4 r0, t0, t1                   "
               
 Samples(24) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Blend D2             " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add_d2 r0, t0, t1                   "
               
 Samples(25) = ";-----------------------------------" & vbNewLine & _
               "; Glow Mapping Blend SAT            " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add_sat r0, t0, t1                  "
               
 Samples(26) = ";-----------------------------------" & vbNewLine & _
               "; Detail Mapping                    " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "add r0, t0, t1_bias                 "
               
 Samples(27) = ";-----------------------------------" & vbNewLine & _
               "; Compare                           " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.2                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "tex t2                              " & vbNewLine & _
               "cmp r0, t0, t1, t2                  "
               
 Samples(28) = ";-----------------------------------" & vbNewLine & _
               "; Sun Blow                          " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.2                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "def c0, 0.1, 0.1, 0.1, 0.1          " & vbNewLine & _
               "def c1, 0.8, 0.8, 0.8, 0.8          " & vbNewLine & _
               "def c2, 0.2, 0.2, 0.2, 1.0          " & vbNewLine & _
               "def c3, 0.6, 0.6, 0.6, 1.0          " & vbNewLine & _
               "def c4, 0.9, 0.9, 0.0, 1.0          " & vbNewLine & _
               "                                    " & vbNewLine & _
               "texcoord t0                         " & vbNewLine & _
               "texcoord t1                         " & vbNewLine & _
               "dp3 r1, t0, t1                      " & vbNewLine & _
               "sub t3, r1, c0                      " & vbNewLine & _
               "cmp_sat r0, t1, c3, c2              " & vbNewLine & _
               "sub t3, r1, c1                      " & vbNewLine & _
               "cmp_sat r0, t3, c4, r0              "
               
 Samples(29) = ";-----------------------------------" & vbNewLine & _
               "; Dark Mapping 1                    " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mad r0, t1, t0, v1                  "
               
 Samples(30) = ";-----------------------------------" & vbNewLine & _
               "; Dark Mapping 2                    " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mul r0, t1, t0                      "
               
 Samples(31) = ";-----------------------------------" & vbNewLine & _
               "; Dark Map Diffuse 1                " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mad r0, t1, t0, v0                  "
               
 Samples(32) = ";-----------------------------------" & vbNewLine & _
               "; Dark Map Diffuse 2                " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "mad_d2 r0, t1_bias, t0_bias, v0     "
               
 Samples(33) = ";-----------------------------------" & vbNewLine & _
               "; Subtract 1                        " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "sub r0, t0, t1                      "
               
 Samples(34) = ";-----------------------------------" & vbNewLine & _
               "; Subtract 2                        " & vbNewLine & _
               ";-----------------------------------" & vbNewLine & _
               "ps.1.1                              " & vbNewLine & _
               "                                    " & vbNewLine & _
               "tex t0                              " & vbNewLine & _
               "tex t1                              " & vbNewLine & _
               "sub r0, t1, t0                      "
              
End Sub

Public Sub ToggleSampleCode(Direction As Direction_)
 Static SamPosition As Long
 
 If SamPosition <= 1 Then frmMain.cmdSamBack.Enabled = False Else frmMain.cmdSamBack.Enabled = True
 If SamPosition >= 33 Then frmMain.cmdSamFor.Enabled = False Else frmMain.cmdSamFor.Enabled = True
 
 Select Case Direction
  Case 0
   SamPosition = SamPosition + 1
   frmMain.txtShader.Text = Samples(SamPosition)
   frmMain.cmdSamBack.Enabled = True
  Case 1
   SamPosition = SamPosition - 1
   frmMain.txtShader.Text = Samples(SamPosition)
   frmMain.cmdSamFor.Enabled = True
 End Select
 
 frmMain.lblSample.Caption = "Sample - " & SamPosition + 1
 
End Sub
