VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3dObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Constants
Private Const PI                As Double = 3.141593
Private Const PENSOLID          As Integer = 0
'Projection variables
Private HCW                     As Long
Private HCH                     As Long
Private HPC                     As Long
Private VPC                     As Long
Private ASPECT_COMP             As Long
'Sine and Cosine lookup tables
Private CosTable(360)           As Single
Private SinTable(360)           As Single
'Our 3d object
Private obj3dObject             As Object3D
'Canvas
Private Render                  As PictureBox

'API type storing and rendering 2d points
Private Type POINTAPI
    X                             As Long
    Y                             As Long
End Type

'type to store a 3d point
Private Type Point3d
    X                             As Single
    Y                             As Single
    Z                             As Single
End Type

'type used to store which points make up a triangle
Private Type Face3D
    A                             As Integer    'Vertex 1 of the triangle
    B                             As Integer    'Vertex 2 of the triangle
    C                             As Integer    'Vertex 3 of the triangle
    Z                             As Integer    'Average distance from viewer of triangle
    N                             As Point3d    'Original normal value
    N2                            As Point3d    'Rotates normal value
    color                         As Long       'color of supplied triangle
End Type

'data type for storing a 3d object
Private Type Object3D
    Name                          As String     'name of the object
    Version                       As String     'version of the object
    NumVertices                   As Integer    'Number of verticies
    NumTriangles                  As Integer    'Number of triangles
    Xangle                        As Long       'X rotation of object
    Yangle                        As Long       'Y rotation of object
    Zangle                        As Long       'Z rotation of object
    ScaleFactor                   As Double     'Scale of object
    CenterofWorld                 As Point3d    'Center of object position
    Light                         As Point3d    'position of light
    LocalCoord()                  As Point3d    'original coordinates of object
    RotatedLocalCoord()           As Point3d    'Rotated coordinates of object
    ScreenCoord()                 As Point3d    'Projected 2d coordinates of object
    Triangle()                    As Face3D     'array to store triangles
    ZOrder                        As Boolean    'wether triangles are painted in order of depth
    Lighted                       As Boolean    'wether light is accounted for in the when rendering
    Style                         As Integer    '1=solid 2=wireframe nontransparent 3=wireframe transparent
End Type

'API Declarations
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
                                              lpPoint As POINTAPI, _
                                              ByVal nCount As Long) As Long

Private Sub CalculateNormals()

  Dim tmpVec1            As Point3d
  Dim tmpVec2            As Point3d
  Dim norm               As Point3d
  Dim lngIncr            As Long

    With obj3dObject
        For lngIncr = 0 To .NumTriangles - 1
            tmpVec1.X = .LocalCoord(.Triangle(lngIncr).C).X - .LocalCoord(.Triangle(lngIncr).B).X
            tmpVec1.Y = .LocalCoord(.Triangle(lngIncr).C).Y - .LocalCoord(.Triangle(lngIncr).B).Y
            tmpVec1.Z = .LocalCoord(.Triangle(lngIncr).C).Z - .LocalCoord(.Triangle(lngIncr).B).Z
            tmpVec2.X = .LocalCoord(.Triangle(lngIncr).A).X - .LocalCoord(.Triangle(lngIncr).B).X
            tmpVec2.Y = .LocalCoord(.Triangle(lngIncr).A).Y - .LocalCoord(.Triangle(lngIncr).B).Y
            tmpVec2.Z = .LocalCoord(.Triangle(lngIncr).A).Z - .LocalCoord(.Triangle(lngIncr).B).Z
            'get the normal of the plane my multiplying the two vectors
            norm = CrossProduct(tmpVec1, tmpVec2)
            'divide the normal by its length to
            .Triangle(lngIncr).N = Normalize(norm)
        Next '  LNGINCR
    End With 'obj3dObject

End Sub

Private Function CrossProduct(V As Point3d, _
                              W As Point3d) As Point3d

    With CrossProduct
        .X = V.Y * W.Z - W.Y * V.Z
        .Y = V.Z * W.X - W.Z * V.X
        .Z = V.X * W.Y - W.X * V.Y
    End With 'CrossProduct

End Function

Private Sub DoRotations()

  Dim lngIncr        As Long
  Dim RotationBuffer As Point3d

    'X rotation error check
    If obj3dObject.Xangle > 360 Then
        obj3dObject.Xangle = (obj3dObject.Xangle Mod 360)
     ElseIf obj3dObject.Xangle < 0 Then 'NOT OBJ3DOBJECT.XANGLE...
        obj3dObject.Xangle = (obj3dObject.Xangle Mod 360 + 360)
    End If
    
    'Y rotation error check
    If obj3dObject.Yangle > 360 Then
        obj3dObject.Yangle = (obj3dObject.Yangle Mod 360)
     ElseIf obj3dObject.Yangle < 0 Then 'NOT OBJ3DOBJECT.YANGLE..'NOT OBJ3DOBJECT.YANGLE...
        obj3dObject.Yangle = (obj3dObject.Yangle Mod 360 + 360)
    End If
    
    'Z rotation error check
    If obj3dObject.Zangle > 360 Then
        obj3dObject.Zangle = (obj3dObject.Zangle Mod 360)
     ElseIf obj3dObject.Zangle < 0 Then 'NOT OBJ3DOBJECT.ZANGLE...
        obj3dObject.Zangle = (obj3dObject.Zangle Mod 360 + 360)
    End If
    
    With obj3dObject
        'rotate local coordinates around each axis
        For lngIncr = 0 To .NumVertices - 1
            'xrotation
            RotationBuffer = .LocalCoord(lngIncr)
            .RotatedLocalCoord(lngIncr).X = .ScaleFactor * (RotationBuffer.X)
            .RotatedLocalCoord(lngIncr).Y = .ScaleFactor * (RotationBuffer.Y * CosTable(.Xangle) - RotationBuffer.Z * SinTable(.Xangle))
            .RotatedLocalCoord(lngIncr).Z = .ScaleFactor * (RotationBuffer.Z * CosTable(.Xangle) + RotationBuffer.Y * SinTable(.Xangle))
            'yrotation
            RotationBuffer = .RotatedLocalCoord(lngIncr)
            .RotatedLocalCoord(lngIncr).X = .ScaleFactor * (RotationBuffer.X * CosTable(.Yangle) + RotationBuffer.Z * SinTable(.Yangle))
            .RotatedLocalCoord(lngIncr).Y = .ScaleFactor * (RotationBuffer.Y)
            .RotatedLocalCoord(lngIncr).Z = .ScaleFactor * (RotationBuffer.Z * CosTable(.Yangle) - RotationBuffer.X * SinTable(.Yangle))
            'zrotation
            RotationBuffer = .RotatedLocalCoord(lngIncr)
            .RotatedLocalCoord(lngIncr).X = .ScaleFactor * (RotationBuffer.X * CosTable(.Zangle) - RotationBuffer.Y * SinTable(.Zangle))
            .RotatedLocalCoord(lngIncr).Y = .ScaleFactor * (RotationBuffer.Y * CosTable(.Zangle) + RotationBuffer.X * SinTable(.Zangle))
            .RotatedLocalCoord(lngIncr).Z = .ScaleFactor * (RotationBuffer.Z)
        Next '  LNGINCR
        
        'rotate normals around each axis
        For lngIncr = 0 To .NumTriangles - 1
            'xrotation
            RotationBuffer = .Triangle(lngIncr).N
            .Triangle(lngIncr).N2.X = (RotationBuffer.X)
            .Triangle(lngIncr).N2.Y = (RotationBuffer.Y * CosTable(.Xangle) - RotationBuffer.Z * SinTable(.Xangle))
            .Triangle(lngIncr).N2.Z = (RotationBuffer.Z * CosTable(.Xangle) + RotationBuffer.Y * SinTable(.Xangle))
            'yrotation
            RotationBuffer = .Triangle(lngIncr).N2
            .Triangle(lngIncr).N2.X = (RotationBuffer.X * CosTable(.Yangle) + RotationBuffer.Z * SinTable(.Yangle))
            .Triangle(lngIncr).N2.Y = (RotationBuffer.Y)
            .Triangle(lngIncr).N2.Z = (RotationBuffer.Z * CosTable(.Yangle) - RotationBuffer.X * SinTable(.Yangle))
            'zrotation
            RotationBuffer = .Triangle(lngIncr).N2
            .Triangle(lngIncr).N2.X = (RotationBuffer.X * CosTable(.Zangle) - RotationBuffer.Y * SinTable(.Zangle))
            .Triangle(lngIncr).N2.Y = (RotationBuffer.Y * CosTable(.Zangle) + RotationBuffer.X * SinTable(.Zangle))
            .Triangle(lngIncr).N2.Z = (RotationBuffer.Z)
        Next '  LNGINCR
    End With 'obj3dObject

End Sub

Private Function DotProduct(V As Point3d, _
                            W As Point3d) As Double

    'Cosine of the angle between two vectors
    DotProduct = (V.X * W.X) + (V.Y * W.Y) + (V.Z * W.Z)

End Function

Public Sub LoadObject(ByVal strFileName As String, _
                      Canvas As PictureBox, _
                      ByVal intStyle As Integer, _
                      ByVal sngCenterofWorldX As Single, _
                      ByVal sngCenterofWorldY As Single, _
                      ByVal sngCenterofWorldZ As Single, _
                      ByVal dblScaleFactor As Double, _
                      ByVal lngSetXRotation As Long, _
                      ByVal lngSetYRotation As Long, _
                      ByVal lngSetZRotation As Long, _
                      ByVal blnZorder As Boolean, _
                      ByVal blnLight As Boolean, _
                      ByVal sngLightX As Single, _
                      ByVal sngLightY As Single, _
                      ByVal sngLightZ As Single)

  Dim i          As Integer
  Dim strTemp    As String

    'get values needed for rendering calculations
    Set Render = Canvas
    HCH = Render.ScaleHeight / 2
    HCW = Render.ScaleWidth / 2
    ASPECT_COMP = (Render.ScaleHeight) / ((Render.ScaleWidth * 3) / 4)
    HPC = HCW / (Tan((60 / 2) * (PI / 180)))
    VPC = HCH / (Tan((60 / 2) * (PI / 180)))
    
    'set initial stuff for object
    With obj3dObject
        .Style = intStyle
        .ZOrder = blnZorder
        .CenterofWorld.X = sngCenterofWorldX
        .CenterofWorld.Y = sngCenterofWorldY
        .CenterofWorld.Z = sngCenterofWorldZ
        .ScaleFactor = dblScaleFactor
        .Xangle = lngSetXRotation
        .Yangle = lngSetYRotation
        .Zangle = lngSetZRotation
        .Lighted = blnLight
        .Light.X = sngLightX
        .Light.Y = sngLightY
        .Light.Z = sngLightZ
    End With 'obj3dObject
    Open strFileName For Input As 1
    
    'make sure file is a valid object file
    Input #1, strTemp
    If strTemp <> "3D OBJECT DEFINITION FILE" Then
        MsgBox "Not a valid object file!", vbOKOnly + vbCritical, "Open"
        Exit Sub
    End If
    
    'get version
    Input #1, strTemp
    obj3dObject.Version = Trim$(strTemp)
    
    'get name
    Input #1, strTemp
    obj3dObject.Name = Trim$(strTemp)
    
    'skip line
    Line Input #1, strTemp
    
    'get number of vertices
    Line Input #1, strTemp
    With obj3dObject
        .NumVertices = Trim$(strTemp)
        ReDim Preserve .LocalCoord(0 To .NumVertices - 1)
        For i = 0 To (.NumVertices - 1)
            'input x,y,z coordinates for vertex
            Input #1, .LocalCoord(i).X, .LocalCoord(i).Y, .LocalCoord(i).Z
        Next i
    End With 'obj3dObject
    
    'skip line
    Line Input #1, strTemp
    
    'get triangles
    Line Input #1, strTemp
    obj3dObject.NumTriangles = Trim$(strTemp)
    With obj3dObject
        ReDim Preserve .Triangle(0 To obj3dObject.NumTriangles - 1)
    End With 'obj3dObject
    For i = 0 To (obj3dObject.NumTriangles - 1)
        Input #1, obj3dObject.Triangle(i).A, obj3dObject.Triangle(i).B, obj3dObject.Triangle(i).C, obj3dObject.Triangle(i).color
    Next i
    Close #1
    
    'redimension arrays
    With obj3dObject
        ReDim Preserve .RotatedLocalCoord(0 To .NumVertices - 1)
        ReDim Preserve .ScreenCoord(0 To .NumVertices - 1)
    End With 'obj3dObject
    
    'set cos and sin values
    MakeTables
    
    'get the normals for the original coordinates
    CalculateNormals

End Sub

Private Sub LocaltoScreen()

  Dim lngIncr   As Long         'counter
  Dim tempcoord As Point3d      'buffer to hold temporary converted coordinates

    With obj3dObject
        For lngIncr = 0 To .NumVertices - 1
            'converts local coordinates into world ones relative to the supplied Canvas
            tempcoord.X = .RotatedLocalCoord(lngIncr).X + .CenterofWorld.X
            tempcoord.Y = .RotatedLocalCoord(lngIncr).Y + .CenterofWorld.Y
            tempcoord.Z = .RotatedLocalCoord(lngIncr).Z + .CenterofWorld.Z
            'now we project the 3d coordinats on to a 2d plane with perspective
            .ScreenCoord(lngIncr).X = (tempcoord.X * HPC / tempcoord.Z) + HCW
            .ScreenCoord(lngIncr).Y = (-tempcoord.Y * VPC * ASPECT_COMP / tempcoord.Z) + HCH
        Next '  LNGINCR
    End With 'obj3dObject

End Sub

Private Function Magnitude(V As Point3d) As Double

    'retrieve the lenth of the vector
    Magnitude = Sqr(V.X * V.X + V.Y * V.Y + V.Z * V.Z)

End Function

Private Sub MakeTables()

  Dim i As Long

    'because the sine and cosine functions are slow we create
    'lookup tables to increase rendering speed
    For i = 0 To 360
        CosTable(i) = Cos((i * PI) / 180)
        SinTable(i) = Sin((i * PI) / 180)
    Next '  I

End Sub

Private Function Normalize(Vect As Point3d) As Point3d

  'converts the vectors length to 1 unit
  
  Dim m As Double

    m = Magnitude(Vect)
    If m = 0 Then
        m = 1
    End If
    With Normalize
        .X = (Vect.X / m)
        .Y = (Vect.Y / m)
        .Z = (Vect.Z / m)
    End With 'Normalize

End Function

Private Sub QuickSortFaces(vntArr() As Face3D, _
                           ByVal lngLeft As Long, _
                           ByVal lngRight As Long)

  Dim i          As Long
  Dim j          As Long
  Dim lngMid     As Long
  Dim vntTestVal As Variant
  Dim vntTemp    As Face3D

    If (lngLeft < lngRight) Then
        lngMid = (lngLeft + lngRight) \ 2
        vntTestVal = vntArr(lngMid).Z
        i = lngLeft
        j = lngRight
        Do
            Do While vntArr(i).Z < vntTestVal
                i = i + 1
            Loop
            Do While vntArr(j).Z > vntTestVal
                j = j - 1
            Loop
            If (i <= j) Then
                vntTemp = vntArr(j)
                vntArr(j) = vntArr(i)
                vntArr(i) = vntTemp
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        If (j <= lngMid) Then
            Call QuickSortFaces(vntArr, lngLeft, j)
            Call QuickSortFaces(vntArr, i, lngRight)
         Else 'NOT (J...
            Call QuickSortFaces(vntArr, i, lngRight)
            Call QuickSortFaces(vntArr, lngLeft, j)
        End If
    End If

End Sub

Private Sub QuickSortZFaces(ByVal NumPoints As Long, _
                            Points2qS() As Point3d, _
                            Faces2qS() As Face3D)

  Dim cnt As Long

    For cnt = 0 To NumPoints
        With Faces2qS(cnt)
            .Z = (Points2qS(.A).Z + Points2qS(.B).Z + Points2qS(.C).Z) \ 3
        End With 'FACES2QS(CNT)
    Next cnt
    Call QuickSortFaces(Faces2qS, 0, NumPoints)

End Sub

Public Sub RenderObject()

  Dim lngIncr              As Long
  Dim ScreenBuffer(0 To 2) As POINTAPI
  Dim Brush                As Long
  Dim Pen                  As Long
  Dim OldBrush             As Long
  Dim OldPen               As Long
  Dim B                    As Single
  Dim CR                   As Byte
  Dim CG                   As Byte
  Dim CB                   As Byte
  Dim color                As Long

    'calculate rotations around all three axis
    DoRotations
    'project our array of points
    LocaltoScreen
    'order triangles by there z values
    If obj3dObject.ZOrder Then
        QuickSortZFaces obj3dObject.NumTriangles - 1, obj3dObject.RotatedLocalCoord, obj3dObject.Triangle
    End If
    'if the style is point then skip rendering polygons
    If obj3dObject.Style = 0 Then
        GoTo POINT
    End If
    For lngIncr = 0 To obj3dObject.NumTriangles - 1
        'if stlye is transparent then skip visible check
        If (obj3dObject.Style = 3) Or (obj3dObject.Style = 4) Then
            GoTo TRANSSTYLE
        End If
        'if face is visible then show it else don't
        If Visible(lngIncr) Then
TRANSSTYLE:
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).A)
                ScreenBuffer(0).X = .X
                ScreenBuffer(0).Y = .Y
            End With 'OBJ3DOBJECT.ROTATEDLOCALCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).A)'OBJ3DOBJECT.SCREENCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).A)
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).B)
                ScreenBuffer(1).X = .X
                ScreenBuffer(1).Y = .Y
            End With 'OBJ3DOBJECT.ROTATEDLOCALCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).B)'OBJ3DOBJECT.SCREENCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).B)
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).C)
                ScreenBuffer(2).X = .X
                ScreenBuffer(2).Y = .Y
            End With 'OBJ3DOBJECT.ROTATEDLOCALCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).C)'OBJ3DOBJECT.SCREENCOORD(OBJ3DOBJECT.TRIANGLE(LNGINCR).C)
            '0 - point
            '1 - solid(culled)
            '2 - wireframe(culled)
            '3 - wireframe
            '4 - solid
            With obj3dObject
                If (.Style = 1) Or (.Style = 4) Then
                    'wether or not we account for light
                    If .Lighted Then
                        .Light = Normalize(.Light)
                        B = DotProduct(.Triangle(lngIncr).N2, .Light)
                        If B < 0 Then
                            B = 0
                        End If
                        color = .Triangle(lngIncr).color
                        CR = color Mod 256
                        CG = color \ 256 Mod 256
                        CB = color \ 65536 Mod 256
                        color = RGB(CR * B, CG * B, CB * B)
                     Else '.LIGHTED = FALSE/0
                        color = .Triangle(lngIncr).color
                    End If
                    Brush = CreateSolidBrush(color)
                    Pen = CreatePen(PENSOLID, 1, color) '.Triangle(lngIncr).Color)
                    With Render
                        OldPen = SelectObject(.hdc, Pen)
                        OldBrush = SelectObject(.hdc, Brush)
                        Polygon .hdc, ScreenBuffer(0), 3
                        SelectObject .hdc, OldPen
                        SelectObject .hdc, OldBrush
                    End With 'Render
                    DeleteObject Pen
                    DeleteObject Brush
                 Else 'NOT (.STYLE...
                    Render.Line (ScreenBuffer(0).X, ScreenBuffer(0).Y)-(ScreenBuffer(1).X, ScreenBuffer(1).Y) ', 2
                    Render.Line (ScreenBuffer(1).X, ScreenBuffer(1).Y)-(ScreenBuffer(2).X, ScreenBuffer(2).Y) ', 2
                    Render.Line (ScreenBuffer(2).X, ScreenBuffer(2).Y)-(ScreenBuffer(0).X, ScreenBuffer(0).Y) ', 2
                End If
            End With 'obj3dObject
        End If
    Next '  LNGINCR

Exit Sub

POINT:
    For lngIncr = 0 To obj3dObject.NumVertices - 1
        Render.PSet (obj3dObject.ScreenCoord(lngIncr).X, obj3dObject.ScreenCoord(lngIncr).Y)
    Next '  LNGINCR

End Sub

Public Property Get RotateX() As Long

    RotateX = obj3dObject.Xangle

End Property

Public Property Get RotateY() As Long

    RotateY = obj3dObject.Yangle

End Property

Public Property Get RotateZ() As Long

    RotateZ = obj3dObject.Zangle

End Property

Public Sub Rotate(Optional ByVal X As Double, _
                        Optional ByVal Y As Double, _
                        Optional ByVal Z As Double)

    If Not (IsMissing(X)) Then
        obj3dObject.Xangle = X
    End If
    If Not (IsMissing(Y)) Then
        obj3dObject.Yangle = Y
    End If
    If Not (IsMissing(Z)) Then
        obj3dObject.Zangle = Z
    End If

End Sub

Public Sub Translate(Optional XPos As Variant, _
                           Optional YPos As Variant, _
                           Optional ZPos As Variant)

    If Not (IsMissing(XPos)) Then
        obj3dObject.CenterofWorld.X = XPos
    End If
    If Not (IsMissing(YPos)) Then
        obj3dObject.CenterofWorld.Y = YPos
    End If
    If Not (IsMissing(ZPos)) Then
        obj3dObject.CenterofWorld.Z = ZPos
    End If

End Sub

Public Property Get TranslateX() As Double

    TranslateX = obj3dObject.CenterofWorld.X

End Property

Public Property Get TranslateY() As Double

    TranslateY = obj3dObject.CenterofWorld.Y

End Property

Public Property Get TranslateZ() As Double

    TranslateZ = obj3dObject.CenterofWorld.Z

End Property

Private Function Visible(ByVal lngFace As Long) As Boolean

  Dim N As Double

    With obj3dObject
        'works more efficiently than comparing the dotproduct of the
        'camera and triangle 3d normals to determine visibility
        
        'What acutally happens here is since the 3d coordinates are now
        'projected onto a 2d plane, we just have to see if the verticies
        'of each triangle are clockwise or counter clockwise. If they are
        'clockwise then the triangle is facing the other way and shouldnt
        'be rendered (if you create your own object file then each of the
        'verticies of the triangles should be made in a clockwise direction
        'when facing forward or it wont be rendered at the right time
        
        N = ((.ScreenCoord(.Triangle(lngFace).A).Y - .ScreenCoord(.Triangle(lngFace).C).Y) * (.ScreenCoord(.Triangle(lngFace).B).X - .ScreenCoord(.Triangle(lngFace).A).X)) - ((.ScreenCoord(.Triangle(lngFace).A).X - .ScreenCoord(.Triangle(lngFace).C).X) * (.ScreenCoord(.Triangle(lngFace).B).Y - .ScreenCoord(.Triangle(lngFace).A).Y))
        If N > 0 Then   'set normal to buffer and set order to renderorder
            Visible = True
         ElseIf N <= 0# Then 'NOT N...
            Visible = False
        End If
    End With 'OBJ3DOBJECT

End Function
