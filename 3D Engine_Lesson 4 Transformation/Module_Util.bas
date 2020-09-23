Attribute VB_Name = "Module_Util"
Option Explicit
'==============================================================================================================
'
'       In this module we define all useful methods
'
'================================================================================================




Function Make_Vertex2D(ByVal Xpos As Single, ByVal Ypos As Single, ByVal Color As Long) As QUEST3D_VERTEXCOLORED2D
With Make_Vertex2D
  .Color = Color
  .Position.z = 0.5
  .rhw = 1
  .Position.x = Xpos
  .Position.y = Ypos



End With

End Function

Function Make_Vertex3D(ByVal Xpos As Single, ByVal Ypos As Single, ByVal Zpos As Single, ByVal Color As Long) As QUEST3D_VERTEXCOLORED3D
With Make_Vertex3D
  .Color = Color
  .Position.z = Zpos
  .Position.x = Xpos
  .Position.y = Ypos



End With

End Function




Function Make_Color(ByVal RedChanel As Byte, ByVal GreenChanel As Byte, ByVal BlueChanel As Byte) As Long
    Make_Color = D3DColorXRGB(RedChanel, GreenChanel, BlueChanel)
End Function


'============================================================================
'
'MATRIX methods
'
'
'=============================================================================
Function Matrix_Get(ByVal Xscal As Single, ByVal Yscal As Single, ByVal Zscal As Single, ByVal Xrot As Single, ByVal Yrot As Single, ByVal Zrot As Single, ByVal Xmov As Single, ByVal Ymov As Single, ByVal Zmov As Single) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_Get = MatZ

End Function


Function Matrix_Ret(MatRet As D3DMATRIX, ByVal Xscal As Single, ByVal Yscal As Single, ByVal Zscal As Single, ByVal Xrot As Single, ByVal Yrot As Single, ByVal Zrot As Single, ByVal Xmov As Single, ByVal Ymov As Single, ByVal Zmov As Single) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Xscal, Yscal, Zscal)
    Call MRotate(ROTz, Xrot, Yrot, Zrot)
    Call D3DXMatrixTranslation(MOVz, Xmov, Ymov, Zmov)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    MatRet = MatZ

End Function

'----------------------------------------
'Name: Matrix_GetEX
'Object: Matrix
'Event: GetEX
'----------------------------------------
'----------------------------------------
'Name: Matrix_GetEX
'Object: Matrix
'Event: GetEX
'Description:
'----------------------------------------
Function Matrix_GetEX(Vscal As D3DVECTOR, vRot As D3DVECTOR, Vtrans As D3DVECTOR) As D3DMATRIX

  Dim MatZ As D3DMATRIX
  Dim ROTz As D3DMATRIX
  Dim MOVz As D3DMATRIX
  Dim TEMPz As D3DMATRIX

    D3DXMatrixIdentity MatZ
    D3DXMatrixIdentity MOVz
    D3DXMatrixIdentity ROTz

    Call D3DXMatrixScaling(MatZ, Vscal.x, Vscal.y, Vscal.z)
    Call MRotate(ROTz, vRot.x, vRot.y, vRot.z)
    Call D3DXMatrixTranslation(MOVz, Vtrans.x, Vtrans.y, Vtrans.z)

    D3DXMatrixMultiply TEMPz, MatZ, ROTz
    D3DXMatrixMultiply MatZ, TEMPz, MOVz

    Matrix_GetEX = MatZ

End Function


Sub MRotate(DestMat As D3DMATRIX, ByVal nValueX As Single, ByVal nValueY As Single, ByVal nValueZ As Single)

  Dim MatX As D3DMATRIX
  Dim MatY As D3DMATRIX
  Dim MatZ As D3DMATRIX
  Dim MatTemp As D3DMATRIX

    D3DXMatrixIdentity MatTemp
    D3DXMatrixIdentity MatX
    D3DXMatrixIdentity MatY
    D3DXMatrixIdentity MatZ

    D3DXMatrixRotationX MatX, nValueX
    D3DXMatrixRotationY MatY, nValueY
    D3DXMatrixRotationZ MatZ, nValueZ

    D3DXMatrixMultiply MatTemp, MatX, MatY
    D3DXMatrixMultiply MatTemp, MatTemp, MatZ

    DestMat = MatTemp

End Sub


'========================================================================
'
'Vector
'====================================================================

Function Vector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR

    
        Vector.x = x
        Vector.y = y
        Vector.z = z

End Function

Function Vector2D(ByVal x As Single, ByVal y As Single) As D3DVECTOR2

    
        Vector2D.x = x
        Vector2D.y = y

End Function
