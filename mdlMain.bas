Attribute VB_Name = "mdlMain"
Option Explicit

'//Realtime raytracer
'//Original (c++) version and other nice
'//Raytrace versions (with shadows, cilinders, etc)
'//Can be found at http://www.2tothex.com/
'//VB port by Almar Joling / quadrantwars@quadrantwars.com
'//Websites: http://www.quadrantwars.com (my game)
'//          http://vbfibre.digitalrice.com (Many VB speed tricks with benchmarks)

'//This code is highly optimized. If you manage to gain some more FPS
'//I'm always interested =-)

'//Finished @ 01/03/2002
'//Feel free to post this code anywhere, but please leave the above info
'//and author info intact. Thank you.

'//Extra optimized by Bill Soo. Thanks!!

Public iFPS As Integer

Private primaryRay As Ray
Private directionTable() As Vector
Private ViewData() As Byte

'//Our light
Public LightLoc As Vector

'//Color
Private Type ColorFloat
    R As Byte
    G As Byte
    B As Byte
End Type

'//Vector
Private Type Vector
    X As Single
    Y As Single
    Z As Single
End Type

'//1 Ray
Private Type Ray
    Origin As Vector
    Direction As Vector
End Type

'//Sphere
Private Type Sphere
    Center As Vector
    Radius As Single
    Color As ColorFloat
    OneOverRadius As Single
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'//Max 3 (or add spheres below)
Public Const numSpheres As Long = 2

'//Similar to refresh
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub Main()

    Dim I As Long
    Dim Spheres(4) As Sphere
    
    '//Quick method..easy to remove as well
    If LCase(Command$) = "uncompiled" Then
        If MsgBox("You are running the raytracer uncompiled. Do you really to continue? " & vbCrLf & "Compiling is recommended!", vbYesNo, "Uncompiled") = vbNo Then
            Unload frmMain
            End
        End If
    End If
    
    MsgBox "If the raytracer gives an error, it's because your monitor is not in 24/32 bit mode...", vbOKOnly, "Info"
    
    '//Allocate the ray direction lookup table
    directionTable = GenerateRayDirectionTable
        
    '//Create a number of spheres
    With Spheres(0)
        .Center.X = 10
        .Center.Y = 100
        .Center.Z = 0
        .Radius = 75
        .Color.R = 255
        .Color.G = 0
        .Color.B = 0
        .OneOverRadius = 1 / .Radius
    End With
    
    With Spheres(1)
        .Center.X = -10
        .Center.Y = -100
        .Center.Z = 20
        .Radius = 50
        .Color.R = 0
        .Color.G = 255
        .Color.B = 0
        .OneOverRadius = 1 / .Radius
    End With
    
    With Spheres(2)
        .Center.X = -100
        .Center.Y = 10
        .Center.Z = 0
        .Radius = 30
        .Color.R = 0
        .Color.G = 132
        .Color.B = 255
        .OneOverRadius = 1 / .Radius
    End With
    
    With Spheres(3)
        .Center.X = 10
        .Center.Y = 100
        .Center.Z = 0
        .Radius = 40
        .Color.R = 255
        .Color.G = 255
        .Color.B = 255
        .OneOverRadius = 1 / .Radius
    End With
    
    With Spheres(3)
        .Center.X = 10
        .Center.Y = 50
        .Center.Z = 40
        .Radius = 9
        .Color.R = 255
        .Color.G = 0
        .Color.B = 255
        .OneOverRadius = 1 / .Radius
    End With
    '//Our position (viewpoint)
    '//Change these values to zoom in, go left/right, etc.
    With primaryRay
        .Origin.X = 0
        .Origin.Y = 0
        .Origin.Z = -600
    End With
    
    '//Light location
    With LightLoc
        .X = 100
        .Y = 100
        .Z = -400
    End With
            
    LoadPicArray2D frmMain.picRay.Picture, ViewSA, ViewBMP, ViewData()

    '//Main loop
    Do
        '// rotate the spheres a bit
        For I = 0 To numSpheres
            Call Rotate(Spheres(I).Center, 0.1 * Sin(10 * I), 0.1 * Sin(10 * I + 2), 0.1 * Sin(10 * I + 1))
        Next I
        Call TraceScene(Spheres, numSpheres, LightLoc)
        '//FPS counter
        iFPS = iFPS + 1
        DoEvents
    Loop
    PicArrayKill ViewData()
End Sub


' //I havent yet bothered to impliment matrix based rotation. this is used so infrequently that it hardly matters though.
Public Sub Rotate(ByRef V As Vector, ByRef ax As Single, ByRef ay As Single, ByRef az As Single)
    Dim Temp As Vector
    Dim sngCosX As Single, sngCosY As Single, sngCosZ As Single
    Dim sngSinX As Single, sngSinY As Single, sngSinZ As Single
    
    '//The less Sin/Cos...the better. Are very slow functions
    '//A lookup table might be used, sacrificing precision
    '//Note: Taylor series do not make it much faster either..
    sngCosX = Cos(ax)
    sngSinX = Sin(ax)
    sngCosY = Cos(ay)
    sngSinY = Sin(ay)
    sngCosZ = Cos(az)
    sngSinZ = Sin(az)
    
    With V
        Temp.Y = .Y
        .Y = (.Y * sngCosX - .Z * sngSinX)
        .Z = (.Z * sngCosX + Temp.Y * sngSinX)
    
        Temp.Z = .Z
        .Z = (.Z * sngCosY - .X * sngSinY)
        .X = (.X * sngCosY + Temp.Z * sngSinY)
    
        Temp.X = .X
        .X = (.X * sngCosZ - .Y * sngSinZ)
        .Y = (.Y * sngCosZ + Temp.X * sngSinZ)
    End With
End Sub

Public Function GenerateRayDirectionTable() As Vector()
    Dim Direction(640& * 480&) As Vector
    Dim currDirection As Vector
    Dim X As Long, Y As Long
    Dim lngPosition As Long
    
    '//Inline should be faster...
    Dim sngScaleFactor As Single
    
    '//Create lookup table...Only used once
    For Y = 0 To 480 - 1
        For X = 0 To 640 - 1
            lngPosition = X + (Y * 640)
            currDirection = Direction(lngPosition)
            currDirection.X = X - 320
            currDirection.Y = Y - 240
            
            '//This value is fairly arbitrary and can basically be interpreted as field of view
            currDirection.Z = 255
            Direction(lngPosition) = currDirection
            
            '// This is definitely not the fastest way to do this. the processor by default computes 1/sqrt and then flips it.
            With Direction(lngPosition)
                sngScaleFactor = 1 / Sqr((.X * .X) + (.Y * .Y) + (.Z * .Z))
                .X = .X * sngScaleFactor
                .Y = .Y * sngScaleFactor
                .Z = .Z * sngScaleFactor
            End With
        Next X
    Next Y
    
    '//Return array
    GenerateRayDirectionTable = Direction
End Function

Public Sub TraceScene(ByRef Spheres() As Sphere, ByRef numSpheres As Integer, ByRef LightLoc As Vector)
    '// setup view rays

    Dim X As Long, Y As Long, Z As Long
    Dim closestIntersectionDistance As Single
    Dim lngBuffer As Long
    Dim bit As RGBQUAD
    Dim rayToSphereCenter As Vector
    Dim lengthRTSC2 As Single
    Dim closestApproach As Single
    Dim halfCord2 As Single
    Dim lngY As Long
    Dim iClosest As Integer
    Dim bHit As Boolean
    Dim sngDistance As Single
    Dim rct As RECT
    Dim lngCalc As Long
    
    '//Changing the loop size will increase FPS very fastly
    '//Especially the outer loop is important!!!
    '//When zooming in, this should be changed to make the tracing 'window'
    '//Larger (but slower!!)
    With rct
        .Top = 90
        .Bottom = 350
        .Left = 260
        .Right = 375
    End With
    
     For Y = 180 To 350
        '//Calculating this only when Y changes should increase a bit...
        lngY = (Y * 640)
        
        For X = 260 To 375
            '//Single dimensional arrays are MUCH faster...
            primaryRay.Direction = directionTable(X + lngY)
                                                                        
                                                                        
            '//an impossibly large value
            closestIntersectionDistance = 1000000
            bHit = False
            '//cycle through all of the spheres to find the closest interesction
            For Z = 0 To numSpheres
                
                    '// this could be optimized for all rays with the same origin (primary and shadow)
                    With rayToSphereCenter
                        .X = Spheres(Z).Center.X - primaryRay.Origin.X
                        .Y = Spheres(Z).Center.Y - primaryRay.Origin.Y
                        .Z = Spheres(Z).Center.Z - primaryRay.Origin.Z
                        closestApproach = (.X * primaryRay.Direction.X) + (.Y * primaryRay.Direction.Y) + (.Z * primaryRay.Direction.Z)
                    End With
                    
                    '//Return false
                    If closestApproach > 0 Then  '// the intersection is behind the ray
                        With rayToSphereCenter
                            lengthRTSC2 = (.X * .X) + (.Y * .Y) + (.Z * .Z) 'length of the ray from the ray's origin to the sphere's center squared
                        End With
                        '//halfCord2 = the distance squared from the closest approach of the ray to a perpendicular to the ray through the center of the sphere to the place where the ray actually intersects the sphere
                        halfCord2 = (Spheres(Z).Radius * Spheres(Z).Radius) - lengthRTSC2 + (closestApproach * closestApproach) '  // sphere.radius * sphere.radius could be precalced, but it might take longer to load it
                                                                                                                                                
                        '//The ray misses the sphere                                                                                                        '// than to calculate it
                        If halfCord2 > 0 Then
                            bHit = True
                            sngDistance = closestApproach - Sqr(halfCord2)
                        
                            If sngDistance < closestIntersectionDistance Then
                                closestIntersectionDistance = sngDistance
                                iClosest = Z
                            End If
                        End If
                    End If
            Next Z
            
            
            '//Something was intersected
            If (closestIntersectionDistance < 1000000) Then
                '//Shade is a pretty big function, that's why it's not inline
                bit = ShadeSphere(Spheres(iClosest), primaryRay, closestIntersectionDistance, LightLoc)
                lngCalc = X * 3
                With bit
                    ViewData(lngCalc, Y) = .rgbBlue
                    ViewData(lngCalc + 1, Y) = .rgbGreen
                    ViewData(lngCalc + 2, Y) = .rgbRed
                End With
            Else '//Make black
                lngCalc = X * 3
                ViewData(lngCalc, Y) = 0
                ViewData(lngCalc + 1, Y) = 0
                ViewData(lngCalc + 2, Y) = 0
            End If
        Next X
    Next Y
    '//"Refresh"
    InvalidateRect frmMain.picRay.hwnd, rct, False
End Sub



Public Function ShadeSphere(ByRef mySphere As Sphere, ByRef myRay As Ray, ByRef Distance As Single, ByRef LightLoc As Vector) As RGBQUAD
    Dim Intersection As Vector
    Dim Normal As Vector
    Dim LightDir As Vector
    Dim LightCoef As Single
    Dim OneOverRadius As Single
    
    Dim sngScaleFactor As Single
    
    '// calculate the location of the intersection between the sphere and the ray.
    With myRay
        Intersection.X = .Origin.X + Distance * .Direction.X
        Intersection.Y = .Origin.Y + Distance * .Direction.Y
        Intersection.Z = .Origin.Z + Distance * .Direction.Z
    End With
    
     '// calculate the normal of the sphere at the point of interesction
    With mySphere
        '// same as ( intersection.x - sphere.center.x ) / sphere.radius
        Normal.X = (Intersection.X - .Center.X) * .OneOverRadius
        Normal.Y = (Intersection.Y - .Center.Y) * .OneOverRadius
        Normal.Z = (Intersection.Z - .Center.Z) * .OneOverRadius
    
        '//Calculate direction from the intersection to light
        
        '//Inline should be faster...
        With LightDir
            .X = LightLoc.X - Intersection.X
            .Y = LightLoc.Y - Intersection.Y
            .Z = LightLoc.Z - Intersection.Z
        '// This is definitely not the fastest way to do this. the processor by default computes 1/sqrt and then flips it.
            sngScaleFactor = 1 / Sqr((.X * .X) + (.Y * .Y) + (.Z * .Z))
        '//Calculate the light coefficient- the value by which the color should be multiplied
            LightCoef = ((Normal.X * .X) + (Normal.Y * .Y) + (Normal.Z * .Z)) * sngScaleFactor
        End With
        
        If (LightCoef < 0) Then LightCoef = 0
        '//Calculate the color to return
        ShadeSphere.rgbRed = .Color.R * LightCoef
        ShadeSphere.rgbGreen = .Color.G * LightCoef
        ShadeSphere.rgbBlue = .Color.B * LightCoef
        
    End With
End Function
