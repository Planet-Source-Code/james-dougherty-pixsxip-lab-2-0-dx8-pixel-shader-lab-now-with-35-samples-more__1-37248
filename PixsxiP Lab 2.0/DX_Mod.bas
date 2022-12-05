Attribute VB_Name = "DX_Mod"
Option Explicit
'This is the same as all DX app's but theres no render loop
'It only renders on Form_Load and the "Compile" button

Public Sub Initialize()
 
 frmMain.Show
 Initialize_DX frmMain.Canvas.hWnd
 InitializeMesh
 UpdateShader
 DoEvents
 Render
 
End Sub

Public Sub Initialize_DX(hWnd As Long)
 '
 'If you get an error look at the Microsoft Direct3D examples
 'On your computer , and see what setup works for yours.
 'It's the D3DPP causing the picture to stay black, see what works for yours
 '
 On Local Error Resume Next
 Dim Mode As D3DDISPLAYMODE
 Dim Caps As D3DCAPS8
    
 If DX8 Is Nothing Then Set DX8 = New DirectX8
 If D3DX Is Nothing Then Set D3DX = New D3DX8
 If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
 Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, Mode)
  
 D3DPP.Windowed = 1
 D3DPP.SwapEffect = D3DSWAPEFFECT_COPY
 'alternative - D3DSWAPEFFECT_DISCARD
 D3DPP.BackBufferFormat = Mode.Format
 D3DPP.BackBufferCount = 1
 D3DPP.hDeviceWindow = hWnd
 
 D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
 MaxVideoCardTextures = Caps.MaxSimultaneousTextures
 Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
 
 D3DD.SetRenderState D3DRS_LIGHTING, 0
 D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
End Sub

Public Sub Render()
 Dim i As Long
 
 D3DD.BeginScene
 D3DD.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0, 0, 0
  
 D3DD.SetPixelShader PS_Handle
 For i = 0 To MaxVideoCardTextures - 1
  D3DD.SetTexture i, Mesh.Texture(i)
 Next
 D3DD.DrawPrimitive D3DPT_TRIANGLEFAN, 0, 2
                                    
 D3DD.EndScene
 D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
 DoEvents
  
End Sub

Public Sub Cleanup_DX8()
 On Local Error Resume Next
 Dim i As Long
 
 If PS_Handle Then
  Call D3DD.DeletePixelShader(PS_Handle)
  PS_Handle = 0
 End If
    
 For i = 0 To 5
  If Not Mesh.Texture(i) Is Nothing Then Set Mesh.Texture(i) = Nothing
 Next
    
 If Not DX_VB Is Nothing Then Set DX_VB = Nothing
 If Not D3DD Is Nothing Then Set D3DD = Nothing
 If Not D3D Is Nothing Then Set D3D = Nothing
 If Not D3DX Is Nothing Then Set D3DX = Nothing
 If Not DX8 Is Nothing Then Set DX8 = Nothing
End Sub
