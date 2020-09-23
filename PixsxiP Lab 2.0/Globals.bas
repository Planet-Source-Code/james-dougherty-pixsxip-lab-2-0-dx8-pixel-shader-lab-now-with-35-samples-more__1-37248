Attribute VB_Name = "Globals"
Option Explicit

Public Type ShaderVertex__1
 X As Single
 Y As Single
 Z As Single
 RHW As Single
 Color As Long
 Color1 As Long
 tU As Single
 TV As Single
 tU1 As Single
 tV1 As Single
End Type

Public Type ShaderObject__1
 Vertices(3) As ShaderVertex__1
 Texture(5) As Direct3DTexture8
End Type

Public Enum Direction_
 Forward = 0
 Backward = 1
End Enum

Public DX8 As New DirectX8
Public D3D As Direct3D8
Public D3DD As Direct3DDevice8
Public D3DX As New D3DX8
Public D3DPP As D3DPRESENT_PARAMETERS
Public DX_VB As Direct3DVertexBuffer8
Public MaxVideoCardTextures As Long
Public HasTexture1 As Boolean

Public Mesh As ShaderObject__1
Public PixelShader As String
Public Samples(50) As String
Public PS_Handle As Long
Public PS_Constants(3, 1) As Single

Public Const FVF_ShaderVertex = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2)
