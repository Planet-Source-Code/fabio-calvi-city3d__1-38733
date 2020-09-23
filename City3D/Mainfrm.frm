VERSION 5.00
Begin VB.Form Mainfrm 
   AutoRedraw      =   -1  'True
   Caption         =   "3D City RM"
   ClientHeight    =   4815
   ClientLeft      =   -1755
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "LOADING..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Main Code was written by K.Sudhakar(sudhakarj21@yahoo.com)
'Update and additions by Fabio Calvi(fabiocalvi@yahoo.com)
Option Explicit

Private WithEvents RM_E As BDirectx
Attribute RM_E.VB_VarHelpID = -1

Dim RMC As New BDirectx
Dim running As Boolean
Dim night As Boolean

Dim mat As Direct3DRMMaterial2

Dim BuildTexture(5)  As Direct3DRMTexture3
Dim GroundTexture As Direct3DRMTexture3
Dim FloorTexture As Direct3DRMTexture3
Dim Septexture As Direct3DRMTexture3
Dim Citytexture As Direct3DRMTexture3
Dim WP(3) As Direct3DRMTexture3
Dim RoofTexture As Direct3DRMTexture3

Dim Mycar As Direct3DRMFrame3
Dim car(4) As Direct3DRMFrame3
Dim CARANIM As Direct3DRMAnimation2
Dim Heli(2) As Direct3DRMFrame3
Dim Helipos(2) As D3DVECTOR
Dim Helilook(2) As Direct3DRMFrame3

Dim Mycarpos As D3DVECTOR
Dim carpos(4) As D3DVECTOR

Dim CM As Single
Dim CarM As Single
Dim oldt As Single
Dim newt As Single
Dim seenspeed As Single
Dim carSpeed As Single

Dim ViewMode As Integer

Dim Mycarsnd As DirectSoundBuffer
Dim Mycarsnd3D As DirectSound3DBuffer

Dim Carsnd(4) As DirectSoundBuffer
Dim Carsnd3D(4) As DirectSound3DBuffer

Dim Helisnd(2) As DirectSoundBuffer
Dim Helisnd3D(2) As DirectSound3DBuffer

Dim Windsnd As DirectSoundBuffer
Dim Windsnd3D As DirectSound3DBuffer

Dim Tracksnd As DirectSoundBuffer
Dim Tracksnd3d As DirectSound3DBuffer

Dim citywrp As Direct3DRMWrap
Dim cityfrm As Direct3DRMFrame3

Dim DINPUT As DirectInput
Dim DIdevice As DirectInputDevice

Dim at As Single
Dim ax As Single
Dim az As Single
Dim ay As Single
Dim aq As D3DRMQUATERNION

Dim turn(4) As Byte
Dim pi
Dim pressed As Byte
Dim dir As D3DVECTOR
Dim mycardir As D3DVECTOR

Dim up As D3DVECTOR
Dim pos As D3DVECTOR
Dim keyb As DIKEYBOARDSTATE
Dim dll As Single
Dim mycardll As Single
Dim pedal As Boolean
Dim Sfondo As Direct3DRMTexture3
Dim incr
Public cockpit As Boolean
Dim wview, hview
Dim carlooked
Dim carlookedpos As D3DVECTOR
Dim carnum As Byte
Dim I_oD3MeshLamps1 As Direct3DRMMeshBuilder3
Dim I_oD3Light As Direct3DRMLight
Dim I_oD3Frame As Direct3DRMFrame3
Dim music As Boolean
Private Sub get_time()
    If night = True Then
        Set Sfondo = Nothing
        RMC.SceneFrame.SetSceneBackgroundImage Sfondo
        RMC.SceneFrame.SetSceneBackground RGB(0, 0, 0)
        RMC.SceneFrame.SetSceneFogColor RGB(0, 0, 0)
        RMC.DirLight.SetType D3DRMLIGHT_SPOT
        RMC.AmbientLight.SetColorRGB 0, 0, 0
    Else
        Set Sfondo = RMC.D3DRM.LoadTexture("cloud3.bmp")
        RMC.SceneFrame.SetSceneBackgroundImage Sfondo
        RMC.SceneFrame.SetSceneBackground RGB(255, 235, 235)
        RMC.AmbientLight.SetColorRGB 0.2, 0.2, 0.2
        RMC.DirLight.SetType D3DRMLIGHT_DIRECTIONAL
        RMC.DirLightFrame.SetPosition Nothing, 5, 5, -5
        RMC.DirLightFrame.LookAt RMC.SceneFrame, Nothing, 0
        RMC.SceneFrame.SetSceneFogColor RGB(255, 235, 235)
    End If
    'Not ready, yet
    'createpotlight -2, 2, 0.2, 0
    'createpotlight -2, 27, 0.2, 0
    'createpotlight -2, 53, 0.2, 0
    'createpotlight 2, 14, -0.2, 0
    'createpotlight 2, 40, -0.2, 0
    'createpotlight 2, 70, -0.2, 0
    'createpotlight -2, -27, 0.2, 0
    'createpotlight -2, -53, 0.2, 0
    'createpotlight 2, -14, -0.2, 0
    'createpotlight 2, -40, -0.2, 0
    'createpotlight 2, -70, -0.2, 0
    ViewMode = pressed
End Sub
Private Sub Form_Load()
Dim i As Byte
    Me.Show
    DoEvents
    For i = 1 To 4
        turn(i) = 0
    Next
    wview = 5
    hview = 1
    
    pi = 3.14159265358979
    
    Set RM_E = RMC
    RMC.hwnd = Pic.hwnd
    RMC.UseBackBuffer = True
    RMC.Use3DHardware = True
    RMC.StartWindowed
    InitDeviceobjects
    inittexture
    INITGEOMENTRY
    InitSounds
    Mycarsnd.Play DSBPLAY_LOOPING
    Carsnd(1).Play DSBPLAY_LOOPING
    Carsnd(2).Play DSBPLAY_LOOPING
    Carsnd(3).Play DSBPLAY_LOOPING
    Carsnd(4).Play DSBPLAY_LOOPING
    Helisnd(1).Play DSBPLAY_LOOPING
    Helisnd(0).Play DSBPLAY_LOOPING
    Windsnd.Play DSBPLAY_LOOPING
    Set DINPUT = RMC.dX.DirectInputCreate
    Set DIdevice = DINPUT.CreateDevice("GUID_SysKeyboard")
    DIdevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIdevice.SetCooperativeLevel Me.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
        
    Set Sfondo = RMC.D3DRM.LoadTexture("cloud3.bmp")
    RMC.SceneFrame.SetSceneBackgroundImage Sfondo
        
    incr = 0
    running = True
    RMC.createhud
    
    Do While running = True
        newt = RMC.dX.TickCount
        seenspeed = (newt - oldt) / 400
        carSpeed = (newt - oldt) / 50
        SceneMove
        RMC.Render
        oldt = newt
        If music = True Then Tracksnd.Play DSBPLAY_LOOPING Else Tracksnd.Stop
        DoEvents
    Loop
End Sub
Private Sub SceneMove()
            
    DIdevice.Acquire
    DIdevice.GetDeviceStateKeyboard keyb
     
    If keyb.Key(DIK_1) Then carnum = 1
    If keyb.Key(DIK_2) Then carnum = 2
    If keyb.Key(DIK_3) Then carnum = 3
    If keyb.Key(DIK_4) Then carnum = 4

    If seenspeed < 0 Then seenspeed = 0
    If seenspeed > 1 Then seenspeed = 1
    
    CM = CM + seenspeed
    
    If CM >= 360 Then CM = 0
    
    ''Car
    If carSpeed < 0 Then carSpeed = 0
    If carSpeed > 1 Then carSpeed = 1
    
    CarM = CarM + carSpeed
   
    If CarM >= 360 Then CarM = 0
    
    'Car 1 movement
    ''''''''''''''''''''''''''''''''''''''''''''''
    car(1).GetPosition Nothing, carpos(1)
    Carsnd3D(1).SetPosition carpos(1).x, carpos(1).y, carpos(1).Z, DS3D_IMMEDIATE
    If turn(1) = 0 Then car(1).SetPosition Nothing, carpos(1).x, carpos(1).y, carpos(1).Z - 0.05
    If carpos(1).Z < -56 And turn(1) = 0 Then
       car(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       turn(1) = 1
       carpos(1).Z = -55.5
    End If
    If turn(1) = 1 Then car(1).SetPosition Nothing, carpos(1).x + 0.05, carpos(1).y, carpos(1).Z
    If carpos(1).x > 56 And turn(1) = 1 Then
       car(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       turn(1) = 2
       carpos(1).x = 55.5
    End If
    If turn(1) = 2 Then car(1).SetPosition Nothing, carpos(1).x, carpos(1).y, carpos(1).Z + 0.05
    If carpos(1).Z > 56 And turn(1) = 2 Then
       car(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       turn(1) = 3
       carpos(1).Z = 55.5
    End If
    If turn(1) = 3 Then car(1).SetPosition Nothing, carpos(1).x - 0.05, carpos(1).y, carpos(1).Z
    If carpos(1).x < -1.2 And turn(1) = 3 Then
       car(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       turn(1) = 0
       carpos(1).x = -1.1
    End If
    
    'Car 2 movement
    ''''''''''''''''''''''''''''''''''''''''''''''
    car(2).GetPosition Nothing, carpos(2)
    Carsnd3D(2).SetPosition carpos(2).x, carpos(2).y, carpos(2).Z, DS3D_IMMEDIATE
    If turn(2) = 0 Then car(2).SetPosition Nothing, carpos(2).x, carpos(2).y, carpos(2).Z + 0.05
    If carpos(2).Z > 54 And turn(2) = 0 Then
       car(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(2) = 1
       carpos(2).Z = 54
    End If
    If turn(2) = 1 Then car(2).SetPosition Nothing, carpos(2).x + 0.05, carpos(2).y, carpos(2).Z
    If carpos(2).x > 54 And turn(2) = 1 Then
       car(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(2) = 2
       carpos(2).x = 54
    End If
    If turn(2) = 2 Then car(2).SetPosition Nothing, carpos(2).x, carpos(2).y, carpos(2).Z - 0.05
    If carpos(2).Z < 1.2 And turn(2) = 2 Then
       car(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(2) = 3
       carpos(2).Z = 1.1
    End If
    If turn(2) = 3 Then car(2).SetPosition Nothing, carpos(2).x - 0.05, carpos(2).y, carpos(2).Z
    If carpos(2).x < -54 And turn(2) = 3 Then
       car(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(2) = 0
       carpos(2).x = -54
    End If
    
    'Car 3 movement
    ''''''''''''''''''''''''''''''''''''''''''''
    car(3).GetPosition Nothing, carpos(3)
    Carsnd3D(3).SetPosition carpos(3).x, carpos(3).y, carpos(3).Z, DS3D_IMMEDIATE
    If turn(3) = 0 Then car(3).SetPosition Nothing, carpos(3).x, carpos(3).y, carpos(3).Z + 0.05
    If carpos(3).Z > -1.2 And turn(3) = 0 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(3) = 1
       carpos(3).Z = -1.1
    End If
    If turn(3) = 1 Then car(3).SetPosition Nothing, carpos(3).x + 0.05, carpos(3).y, carpos(3).Z
    If carpos(3).x > 1.2 And turn(3) = 1 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       turn(3) = 2
       carpos(3).x = 1.1
    End If
    If turn(3) = 2 Then car(3).SetPosition Nothing, carpos(3).x, carpos(3).y, carpos(3).Z + 0.05
    If carpos(3).Z > 54 And turn(3) = 2 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(3) = 3
       carpos(3).Z = 54
    End If
    If turn(3) = 3 Then car(3).SetPosition Nothing, carpos(3).x + 0.05, carpos(3).y, carpos(3).Z
    If carpos(3).x > 54 And turn(3) = 3 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(3) = 4
       carpos(3).x = 54
    End If
    If turn(3) = 4 Then car(3).SetPosition Nothing, carpos(3).x, carpos(3).y, carpos(3).Z - 0.05
    If carpos(3).Z < -54 And turn(3) = 4 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(3) = 5
       carpos(3).Z = -54
    End If
    If turn(3) = 5 Then car(3).SetPosition Nothing, carpos(3).x - 0.05, carpos(3).y, carpos(3).Z
    If carpos(3).x < -54 And turn(3) = 5 Then
       car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(3) = 0
       carpos(3).x = -54
    End If
   
    'Car 4 movement
    ''''''''''''''''''''''''''''''''''''''''''''''
    car(4).GetPosition Nothing, carpos(4)
    Carsnd3D(4).SetPosition carpos(4).x, carpos(4).y, carpos(4).Z, DS3D_IMMEDIATE
    If turn(4) = 0 Then car(4).SetPosition Nothing, carpos(4).x - 0.05, carpos(4).y, carpos(4).Z
    If carpos(4).x < -54 And turn(4) = 0 Then
       car(4).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(4) = 1
       carpos(4).x = -54
    End If
    If turn(4) = 1 Then car(4).SetPosition Nothing, carpos(4).x, carpos(4).y, carpos(4).Z + 0.05
    If carpos(4).Z > 54 And turn(4) = 1 Then
       car(4).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(4) = 2
       carpos(4).Z = 54
    End If
    If turn(4) = 2 Then car(4).SetPosition Nothing, carpos(4).x + 0.05, carpos(4).y, carpos(4).Z
    If carpos(4).x > 54 And turn(4) = 2 Then
       car(4).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(4) = 3
       carpos(4).x = 54
    End If
    If turn(4) = 3 Then car(4).SetPosition Nothing, carpos(4).x, carpos(4).y, carpos(4).Z - 0.05
    If carpos(4).Z < -54 And turn(4) = 3 Then
       car(4).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       turn(4) = 0
       carpos(4).Z = -54
    End If
   
    ''Helicopter1
    Helilook(0).SetPosition Nothing, 50 * Sin((CM - 20) / 10), 10 + (5 * Cos((CM + 20) / 10)), 30 * Cos((CM - 20) / 10)
    Heli(0).SetPosition Nothing, 55 * Sin(CM / 10), 10 + (5 * Cos(CM / 10)), 35 * Cos(CM / 10)
    Heli(0).GetPosition Nothing, Helipos(0)
    Helisnd3D(0).SetPosition Helipos(0).x, Helipos(0).y, Helipos(0).Z, DS3D_IMMEDIATE
    
    ''Helicopter2
    Helilook(1).SetPosition Nothing, 32 * Sin((CM - 1) / 10), Helipos(1).y + Cos(CM / 10), 32 * Cos((CM - 1) / 10)
    Heli(1).SetPosition Nothing, 30 * Sin(CM / 10), Helipos(1).y, 30 * Cos(CM / 10)
    Heli(1).GetPosition Nothing, Helipos(1)
    Helisnd3D(1).SetPosition Helipos(1).x, Helipos(1).y, Helipos(1).Z, DS3D_IMMEDIATE
    
    cockpit = False
     
     If carnum = 1 Then
        carlooked = car(1)
        carlookedpos.x = carpos(1).x
        carlookedpos.y = carpos(1).y
        carlookedpos.Z = carpos(1).Z
     End If
     If carnum = 2 Then
        carlooked = car(2)
        carlookedpos.x = carpos(2).x
        carlookedpos.y = carpos(2).y
        carlookedpos.Z = carpos(2).Z
     End If
     If carnum = 3 Then
        carlooked = car(3)
        carlookedpos.x = carpos(3).x
        carlookedpos.y = carpos(3).y
        carlookedpos.Z = carpos(3).Z
     End If
     If carnum = 4 Then
        carlooked = car(4)
        carlookedpos.x = carpos(4).x
        carlookedpos.y = carpos(4).y
        carlookedpos.Z = carpos(4).Z
     End If
   
    ''Camera views
    Select Case ViewMode
        Case 0
            RMC.CameraFrame.SetPosition Nothing, 5 * Sin(CM / 10), 0.5 + (14 * Abs(Sin(CM / 10))), 4 * Cos(CM / 10)
            RMC.CameraFrame.LookAt carlooked, Nothing, D3DRMCONSTRAIN_Z
            pressed = 0
      
        Case 1
            Static Lastcm As Single
            Static f As Boolean
            RMC.CameraFrame.LookAt carlooked, Nothing, D3DRMCONSTRAIN_Z
                RMC.CameraFrame.SetPosition Nothing, 57 * Sin(CM / 10), 1, 1 * Cos(CM / 10)
            pressed = 1
        
        Case 2
            RMC.CameraFrame.SetPosition Nothing, Helipos(1).x + (5 * Cos(CM / 10)), Helipos(1).y + 1, Helipos(1).Z + (5 * Sin(CM / 10))
            RMC.CameraFrame.LookAt carlooked, Nothing, D3DRMCONSTRAIN_Z
            pressed = 2
        
        Case 3
            RMC.CameraFrame.LookAt carlooked, Nothing, D3DRMCONSTRAIN_Z
            RMC.CameraFrame.SetPosition Nothing, carlookedpos.x + (1.1 * Sin(CM / 20)), 0.5 + (3 * Abs(Sin(CM / 40))), carlookedpos.Z + (4 * Cos(CM / 20))  ' 50 * Sin(CM / 10), 1 + 5 + (5 * Cos(CM / 20)), 33 * Cos(CM / 10)
            pressed = 3
        
        Case 4
            RMC.CameraFrame.SetPosition Nothing, Helipos(0).x - (4 * Sin(CM / 20)), Helipos(0).y + (7 * Sin(CM / 10)), (4 * Cos(CM / 20)) - Helipos(0).Z
            RMC.CameraFrame.LookAt carlooked, Nothing, D3DRMCONSTRAIN_Z
            pressed = 4
        
        Case 5
            If keyb.Key(DIK_RIGHT) Then
                mycardll = seenspeed
                incr = incr - 0.1
                If incr < -1 Then incr = -1
            Else
                If mycardll > 0 Then
                    mycardll = 0
                End If
                If incr < 0 Then incr = incr + 0.1
            End If
            If keyb.Key(DIK_LEFT) Then
                mycardll = -seenspeed
                incr = incr + 0.1
                If incr > 1 Then incr = 1
            
            Else
                If incr > 0 Then incr = incr - 0.1
                
                If mycardll < 0 Then
                    mycardll = 0
                End If

            End If
            
            Mycar.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, mycardll
           
            Mycar.GetPosition Nothing, Mycarpos
            If keyb.Key(DIK_DOWN) Then
                Mycarpos.x = Mycarpos.x - (dir.x * seenspeed * 7)
                Mycarpos.Z = Mycarpos.Z - (dir.Z * seenspeed * 7)
            End If
            If keyb.Key(DIK_UP) Then
                Mycarpos.x = Mycarpos.x + (dir.x * seenspeed * 7)
                Mycarpos.Z = Mycarpos.Z + (dir.Z * seenspeed * 7)
            End If

            
            Mycar.SetPosition Nothing, Mycarpos.x, Mycarpos.y, Mycarpos.Z
            Mycarsnd3D.SetPosition Mycarpos.x, Mycarpos.y, Mycarpos.Z, DS3D_IMMEDIATE
            
            If keyb.Key(DIK_Z) Then hview = hview + 0.1
            If hview > 1.5 Then hview = 1.5
            If keyb.Key(DIK_X) Then hview = hview - 0.1
            If hview < 0.5 Then hview = 0.5
            If keyb.Key(DIK_Q) Then wview = wview + 0.1
            If wview > 7 Then wview = 7
            If keyb.Key(DIK_W) Then wview = wview - 0.1
            If wview < 5 Then wview = 5
            
            RMC.CameraFrame.SetPosition Mycar, wview, hview, 0 - incr
            RMC.CameraFrame.LookAt Mycar, Nothing, D3DRMCONSTRAIN_Z
            
            pressed = 5
        
        Case 6
            cockpit = True
            
            DIdevice.Acquire
            DIdevice.GetDeviceStateKeyboard keyb
            If keyb.Key(DIK_RIGHT) Then
                If pedal = True Then dll = seenspeed
            Else
                If dll > 0# Then
                    dll = 0
                End If
            End If
            If keyb.Key(DIK_LEFT) Then
               If pedal = True Then dll = -seenspeed
            Else
                If dll < 0# Then
                    dll = 0
                End If
            End If
            RMC.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, dll
            RMC.CameraFrame.GetOrientation Nothing, dir, up
            RMC.CameraFrame.GetPosition Nothing, pos
            If keyb.Key(DIK_UP) Then
                pos.x = pos.x + (dir.x * seenspeed * 5)
                pos.Z = pos.Z + (dir.Z * seenspeed * 5)
                pedal = True
            ElseIf keyb.Key(DIK_DOWN) Then
                pos.x = pos.x - (dir.x * seenspeed * 5)
                pos.Z = pos.Z - (dir.Z * seenspeed * 5)
                pedal = True
            Else
                pedal = False
            End If
            
             RMC.CameraFrame.SetPosition Nothing, pos.x, pos.y, pos.Z
             Mycarsnd3D.SetPosition pos.x, pos.y, pos.Z, DS3D_IMMEDIATE
           
            pressed = 6

        Case 11
            If night = True Then night = False Else night = True
            get_time

    End Select

    RMC.CameraFrame.GetPosition Nothing, pos
    RMC.CameraFrame.GetOrientation Nothing, dir, up
    RMC.DsoundLis70.SetPosition pos.x, pos.y, pos.Z, DS3D_IMMEDIATE
    RMC.DsoundLis70.SetOrientation dir.x, dir.y, dir.Z, up.x, up.y, up.Z, DS3D_IMMEDIATE
    
    If night = True Then
        RMC.DirLightFrame.SetPosition Nothing, pos.x, pos.y, pos.Z
        RMC.DirLightFrame.SetOrientation Nothing, dir.x, dir.y, dir.Z, up.x, up.y, up.Z
    End If
    
End Sub
Private Sub Cleanup()
    Set BuildTexture(0) = Nothing
    Set BuildTexture(1) = Nothing
    Set BuildTexture(2) = Nothing
    Set BuildTexture(3) = Nothing
    Set BuildTexture(4) = Nothing
    Set BuildTexture(5) = Nothing
    Set WP(0) = Nothing
    Set WP(1) = Nothing
    Set WP(2) = Nothing
    Set WP(3) = Nothing
    Set FloorTexture = Nothing
    Set RoofTexture = Nothing
    Set Septexture = Nothing
    
    Set Mycar = Nothing
    Set car(1) = Nothing
    Set car(2) = Nothing
    Set car(3) = Nothing
    Set car(4) = Nothing
    Set Heli(1) = Nothing
    Set Heli(0) = Nothing
    Set Helilook(1) = Nothing
    Set Helilook(0) = Nothing
    
    Set Mycarsnd = Nothing
    Set Carsnd(1) = Nothing
    Set Carsnd(2) = Nothing
    Set Carsnd(3) = Nothing
    Set Carsnd(4) = Nothing
    Set Helisnd(1) = Nothing
    Set Helisnd(0) = Nothing
    Set Tracksnd = Nothing
    Set Windsnd = Nothing
    Set Mycarsnd3D = Nothing
    Set Carsnd3D(1) = Nothing
    Set Carsnd3D(2) = Nothing
    Set Carsnd3D(3) = Nothing
    Set Carsnd3D(4) = Nothing
    Set Helisnd3D(1) = Nothing
    Set Helisnd3D(0) = Nothing
    Set Tracksnd3d = Nothing
    Set Windsnd3D = Nothing
    
    Set citywrp = Nothing
    Set cityfrm = Nothing
    Set Citytexture = Nothing
    
    Set RM_E = Nothing
    Set mat = Nothing
    
    Set DINPUT = Nothing
    Set DIdevice = Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    running = False
    Cleanup
End Sub
Private Sub Form_Resize()
    Pic.Width = Me.ScaleWidth
    Pic.Height = Me.ScaleHeight
    If running = False Then Exit Sub
    If RMC.IsFullScreen = True Then Exit Sub
    RMC.Resize Pic.ScaleWidth, Pic.ScaleHeight
End Sub
Private Sub InitDeviceobjects()
    Dim vp As Direct3DRMViewport2
    Set vp = RMC.Viewport
    vp.SetBack 110
    vp.SetFront 1
    vp.SetProjection D3DRMPROJECT_PERSPECTIVE
    
    Set mat = RMC.D3DRM.CreateMaterial(0)
    With mat
        .SetAmbient 1, 1, 1
    End With
    
    RMC.SceneFrame.SetSceneBackground RGB(255, 235, 235)
    RMC.SceneFrame.SetSceneFogEnable D_TRUE
    RMC.SceneFrame.SetSceneFogMethod D3DRMFOGMETHOD_VERTEX
    RMC.SceneFrame.SetSceneFogColor RGB(255, 235, 235)
    RMC.SceneFrame.SetSceneFogMode D3DRMFOG_LINEAR
    'RMC.SceneFrame.SetSceneFogParams 1, 70, 0
    With RMC.Device
        .SetTextureQuality D3DRMTEXTURE_LINEARMIPLINEAR
        .SetQuality D3DRMFILL_SOLID Or D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD Or D3DRMSHADE_GOURAUD
        .SetDither D_TRUE
        .SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY Or D3DRMRENDERMODE_SORTEDTRANSPARENCY
    End With
    
End Sub
Private Sub InitSounds()
Dim i As Byte

    RMC.InitDsound70
    
    Set Mycarsnd = RMC.Create2DsBuffromfile70("engine6.wav")
    Set Mycarsnd3D = RMC.Create3DSBUFfrom2Dbuf70(Mycarsnd)
   
    For i = 1 To 4
        Set Carsnd(i) = RMC.Create2DsBuffromfile70("ENG.wav")
        Set Carsnd3D(i) = RMC.Create3DSBUFfrom2Dbuf70(Carsnd(i))
        Carsnd3D(i).SetMinDistance 2, DS3D_IMMEDIATE
        Carsnd3D(i).SetMaxDistance 1000, DS3D_IMMEDIATE
    Next
    
    Set Helisnd(1) = RMC.Create2DsBuffromfile70("Heli.wav")
    Set Helisnd3D(1) = RMC.Create3DSBUFfrom2Dbuf70(Helisnd(1))
    Set Helisnd(0) = RMC.Create2DsBuffromfile70("Heli.wav")
    Set Helisnd3D(0) = RMC.Create3DSBUFfrom2Dbuf70(Helisnd(0))
    
    Set Windsnd = RMC.Create2DsBuffromfile70("F_wind1.wav")
    Set Windsnd3D = RMC.Create3DSBUFfrom2Dbuf70(Windsnd)
    
    Set Tracksnd = RMC.Create2DsBuffromfile70("Track06.wav")
    Set Tracksnd3d = RMC.Create3DSBUFfrom2Dbuf70(Tracksnd)
    Tracksnd3d.SetMode DS3DMODE_DISABLE, DS3D_IMMEDIATE
    
    Mycarsnd3D.SetMinDistance 1, DS3D_IMMEDIATE
    Mycarsnd3D.SetMaxDistance 10, DS3D_IMMEDIATE
    
    Helisnd3D(1).SetMinDistance 4, DS3D_IMMEDIATE
    Helisnd3D(1).SetMaxDistance 100, DS3D_IMMEDIATE
    Helisnd3D(0).SetMinDistance 4, DS3D_IMMEDIATE
    Helisnd3D(0).SetMaxDistance 100, DS3D_IMMEDIATE
    
    Windsnd3D.SetPosition 55, 55, 55, DS3D_IMMEDIATE
    Windsnd3D.SetMaxDistance 2, DS3D_IMMEDIATE
    
    music = True
End Sub
Private Sub INITGEOMENTRY()
Dim m As Integer
    'Buildings
    CreateBuilding 0, 0, 0, 2, 8, 3
    
    CreateBuilding 20, 0, 1, 2, 8, 2
    CreateBuilding 0, 20, 2, 3, 12, 3
    CreateBuilding -20, 0, 3, 3, 4, 3
    CreateBuilding 0, -20, 4, 3, 4, 3, 6, 6
    
    CreateBuilding 20, 20, 5, 2, 8, 2, 1, 1
    CreateBuilding -20, -20, 0, 3, 4, 3
    CreateBuilding -20, 20, 1, 3, 4, 3
    CreateBuilding 20, -20, 2, 3, 12, 3
    
    CreateBuilding 55, 0, 3, 2, 8, 2
    CreateBuilding 0, 55, 4, 3, 12, 3, 6, 6
    CreateBuilding -55, 0, 5, 3, 4, 3, 1, 1
    CreateBuilding 0, -55, 0, 3, 4, 3
    
    CreateBuilding 55, 55, 1, 2, 8, 2
    CreateBuilding -55, -55, 2, 3, 12, 3
    CreateBuilding -55, 55, 3, 3, 4, 3
    CreateBuilding 55, -55, 4, 3, 4, 3, 6, 6
    
    'Lampspot
    createlightpot -2, 2, 0.2, 0, RoofTexture
    createlightpot -2, 27, 0.2, 0, RoofTexture
    createlightpot -2, 53, 0.2, 0, RoofTexture
    createlightpot 2, 14, -0.2, 0, RoofTexture
    createlightpot 2, 40, -0.2, 0, RoofTexture
    createlightpot 2, 70, -0.2, 0, RoofTexture
    createlightpot -2, -27, 0.2, 0, RoofTexture
    createlightpot -2, -53, 0.2, 0, RoofTexture
    createlightpot 2, -14, -0.2, 0, RoofTexture
    createlightpot 2, -40, -0.2, 0, RoofTexture
    createlightpot 2, -70, -0.2, 0, RoofTexture
    
    createlightpot 27, 53, 0, 0.2, RoofTexture
    createlightpot 53, 53, 0, 0.2, RoofTexture
    createlightpot 82, 53, 0, 0.2, RoofTexture
    createlightpot 14, 57, 0, -0.2, RoofTexture
    createlightpot 40, 57, 0, -0.2, RoofTexture
    createlightpot 68, 57, 0, -0.2, RoofTexture
    createlightpot 27, -53, 0, -0.2, RoofTexture
    createlightpot 53, -53, 0, -0.2, RoofTexture
    createlightpot 82, -53, 0, -0.2, RoofTexture
    createlightpot 14, -57, 0, 0.2, RoofTexture
    createlightpot 40, -57, 0, 0.2, RoofTexture
    createlightpot 68, -57, 0, 0.2, RoofTexture
    
    createlightpot 57, 2, -0.2, 0, RoofTexture
    createlightpot 57, 27, -0.2, 0, RoofTexture
    createlightpot 57, 53, -0.2, 0, RoofTexture
    createlightpot 53, 14, 0.2, 0, RoofTexture
    createlightpot 53, 40, 0.2, 0, RoofTexture
    createlightpot 53, 70, 0.2, 0, RoofTexture
    createlightpot 57, -27, -0.2, 0, RoofTexture
    createlightpot 57, -53, -0.2, 0, RoofTexture
    createlightpot 53, -14, 0.2, 0, RoofTexture
    createlightpot 53, -40, 0.2, 0, RoofTexture
    createlightpot 53, -70, 0.2, 0, RoofTexture
   
    createlightpot -27, 53, 0, 0.2, RoofTexture
    createlightpot -53, 53, 0, 0.2, RoofTexture
    createlightpot -82, 53, 0, 0.2, RoofTexture
    createlightpot -14, 57, 0, -0.2, RoofTexture
    createlightpot -40, 57, 0, -0.2, RoofTexture
    createlightpot -68, 57, 0, -0.2, RoofTexture
    createlightpot -27, -53, 0, -0.2, RoofTexture
    createlightpot -53, -53, 0, -0.2, RoofTexture
    createlightpot -82, -53, 0, -0.2, RoofTexture
    createlightpot -14, -57, 0, 0.2, RoofTexture
    createlightpot -40, -57, 0, 0.2, RoofTexture
    createlightpot -68, -57, 0, 0.2, RoofTexture
    
    createlightpot -57, 2, 0.2, 0, RoofTexture
    createlightpot -57, 27, 0.2, 0, RoofTexture
    createlightpot -57, 53, 0.2, 0, RoofTexture
    createlightpot -53, 14, -0.2, 0, RoofTexture
    createlightpot -53, 40, -0.2, 0, RoofTexture
    createlightpot -53, 70, -0.2, 0, RoofTexture
    createlightpot -57, -27, 0.2, 0, RoofTexture
    createlightpot -57, -53, 0.2, 0, RoofTexture
    createlightpot -53, -14, -0.2, 0, RoofTexture
    createlightpot -53, -40, -0.2, 0, RoofTexture
    createlightpot -53, -70, -0.2, 0, RoofTexture
    
    createlightpot -27, 2, 0, -0.2, RoofTexture
    createlightpot -53, 2, 0, -0.2, RoofTexture
    createlightpot -82, 2, 0, -0.2, RoofTexture
    createlightpot -14, -2, 0, 0.2, RoofTexture
    createlightpot -40, -2, 0, 0.2, RoofTexture
    createlightpot -68, -2, 0, 0.2, RoofTexture
    createlightpot 2, -2, 0, 0.2, RoofTexture
    createlightpot 27, -2, 0, 0.2, RoofTexture
    createlightpot 53, -2, 0, 0.2, RoofTexture
    createlightpot 82, -2, 0, 0.2, RoofTexture
    createlightpot 14, 2, 0, -0.2, RoofTexture
    createlightpot 40, 2, 0, -0.2, RoofTexture
    createlightpot 68, 2, 0, -0.2, RoofTexture
    
    'Ground & roads
    CreatePlaneInScene 0, -0.05, 0, 165, 0, 165, 1, FloorTexture, 1, 1 'because of an awfull visual effect
    
    CreatePlaneInScene 0, 0, 0, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene 55, 0, 0, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene 0, 0, 55, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene -55, 0, 0, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene 0, 0, -55, 55, 0, 55, 1, GroundTexture, 1, 1
    
    CreatePlaneInScene 55, 0, 55, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene -55, 0, -55, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene -55, 0, 55, 55, 0, 55, 1, GroundTexture, 1, 1
    CreatePlaneInScene 55, 0, -55, 55, 0, 55, 1, GroundTexture, 1, 1
    
    'Walls
    Createsep 0, 0
    Createsep 0, 55
    Createsep 55, 0
    Createsep 55, 55
    Createsep 0, -55
    Createsep -55, 0
    Createsep -55, -55
    Createsep -55, 55
    Createsep 55, -55
    
    'Posters
    createwallposter 20, 2.8, WP(0)
    createwallposter 25, 53, WP(1)
    createwallposter -20, -2.8, WP(2)
    createwallposter -25, -53, WP(3)
    
    Dim BOX As D3DRMBOX
    
    'Mycar init
    Dim mymscf As Direct3DRMMeshBuilder3
    Dim mymsch As Direct3DRMMesh
    
    Set mymscf = RMC.D3DRM.CreateMeshBuilder
    mymscf.LoadFromFile "car1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Set mymsch = mymscf.CreateMesh
    Set Mycar = RMC.D3DRM.CreateFrame(Nothing)
    mymsch.ScaleMesh 0.013, 0.01, 0.013

    Mycar.AddVisual mymsch
    RMC.SceneFrame.AddVisual Mycar

    Call Mycar.GetBox(BOX)
    Mycarpos.y = (BOX.Max.y - BOX.min.y)
    Mycarpos.x = (BOX.Max.x - BOX.min.x)
    Mycarpos.Z = (BOX.Max.Z - BOX.min.Z)
    
    Mycar.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
    Mycar.SetPosition Nothing, 54, Mycarpos.y, 70
    
    Dim mscf(4) As Direct3DRMMeshBuilder3
    Dim msch(4) As Direct3DRMMesh
    
    'Car 1 init
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set mscf(1) = RMC.D3DRM.CreateMeshBuilder
    mscf(1).LoadFromFile "car2.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Set msch(1) = mscf(1).CreateMesh
    msch(1).ScaleMesh 0.013, 0.01, 0.013
    Set car(1) = RMC.D3DRM.CreateFrame(Nothing)
    
    car(1).AddVisual msch(1)
    RMC.SceneFrame.AddVisual car(1)
    
    Call car(1).GetBox(BOX)
    carpos(1).y = (BOX.Max.y - BOX.min.y)
    carpos(1).x = (BOX.Max.x - BOX.min.x)
    carpos(1).Z = (BOX.Max.Z - BOX.min.Z)
  
    car(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
    car(1).SetPosition Nothing, -1.1, carpos(1).y, carpos(1).Z
    
    Set mscf(1) = Nothing
    Set msch(1) = Nothing
       
    'Car 2 init
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set mscf(2) = RMC.D3DRM.CreateMeshBuilder
    mscf(2).LoadFromFile "pickup.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Set msch(2) = mscf(2).CreateMesh
    msch(2).ScaleMesh 0.013, 0.01, 0.013
    Set car(2) = RMC.D3DRM.CreateFrame(Nothing)
    
    car(2).AddVisual msch(2)
    RMC.SceneFrame.AddVisual car(2)
    
    Call car(2).GetBox(BOX)
    carpos(2).y = (BOX.Max.y - BOX.min.y)
    carpos(2).x = (BOX.Max.x - BOX.min.x)
    carpos(2).Z = (BOX.Max.Z - BOX.min.Z)
  
    car(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
    car(2).SetPosition Nothing, 1, carpos(2).y, 54
    
    Set mscf(2) = Nothing
    Set msch(2) = Nothing
    
    'Car 3 init
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set mscf(3) = RMC.D3DRM.CreateMeshBuilder
    mscf(3).LoadFromFile "car2.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    For m = 18 To 515
        mscf(3).GetFace(m).SetColorRGB 125, 125, 125
    Next
    For m = 536 To 662
        mscf(3).GetFace(m).SetColorRGB 125, 125, 125
    Next
    For m = 1162 To 1185
        mscf(3).GetFace(m).SetColorRGB 125, 125, 125
    Next
    Set msch(3) = mscf(3).CreateMesh
    msch(3).ScaleMesh 0.013, 0.01, 0.013
    Set car(3) = RMC.D3DRM.CreateFrame(Nothing)
    
    car(3).AddVisual msch(3)
    RMC.SceneFrame.AddVisual car(3)
    
    Call car(3).GetBox(BOX)
    carpos(3).y = (BOX.Max.y - BOX.min.y)
    carpos(3).x = (BOX.Max.x - BOX.min.x)
    carpos(3).Z = (BOX.Max.Z - BOX.min.Z)
  
    car(3).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
    car(3).SetPosition Nothing, -54, carpos(3).y, -54
    
    Set mscf(3) = Nothing
    Set msch(3) = Nothing
    
    'Car 4 init
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set mscf(4) = RMC.D3DRM.CreateMeshBuilder
    mscf(4).LoadFromFile "car2.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    For m = 18 To 515
        mscf(4).GetFace(m).SetColorRGB 0, 210, 210
    Next
    For m = 536 To 662
        mscf(4).GetFace(m).SetColorRGB 0, 210, 210
    Next
    For m = 1162 To 1185
        mscf(4).GetFace(m).SetColorRGB 0, 210, 210
    Next
    Set msch(4) = mscf(4).CreateMesh
    msch(4).ScaleMesh 0.013, 0.01, 0.013
    Set car(4) = RMC.D3DRM.CreateFrame(Nothing)
    
    car(4).AddVisual msch(4)
    RMC.SceneFrame.AddVisual car(4)
    
    Call car(4).GetBox(BOX)
    carpos(4).y = (BOX.Max.y - BOX.min.y)
    carpos(4).x = (BOX.Max.x - BOX.min.x)
    carpos(4).Z = (BOX.Max.Z - BOX.min.Z)
  
    car(4).SetPosition Nothing, 80, carpos(4).y, -54
    
    Set mscf(4) = Nothing
    Set msch(4) = Nothing
    
    'Helicopter1 Geometry
    Set Helilook(1) = RMC.D3DRM.CreateFrame(Nothing)
    Set Heli(1) = RMC.D3DRM.CreateFrame(Nothing)
    Helipos(1).y = 10
    
    'Helicopter2 Geometry
    Set Helilook(0) = RMC.D3DRM.CreateFrame(Nothing)
    Set Heli(0) = RMC.D3DRM.CreateFrame(Nothing)
    Helipos(0).y = 10
    
    
    CreatePlaneInScene 0, 3, 5, 7, 1, 0, 2, Citytexture, 1, 1
    CreatePlaneInScene 0, 3, -5, 7, 1, 0, 2, Citytexture, 1, 1
    CreatePlaneInScene 5, 3, 0, 0, 1, 7, 2, Citytexture, 1, 1
    CreatePlaneInScene -5, 3, 0, 0, 1, 7, 2, Citytexture, 1, 1
    
    CreatePlaneInScene 3, 1.25, 5, 1, 2.5, 0, 2, Citytexture, 1, 1
    CreatePlaneInScene 3, 1.25, -5, 1, 2.5, 0, 2, Citytexture, 1, 1
    CreatePlaneInScene -3, 1.25, 5, 1, 2.5, 0, 2, Citytexture, 1, 1
    CreatePlaneInScene -3, 1.25, -5, 1, 2.5, 0, 2, Citytexture, 1, 1
    
    CreatePlaneInScene 5, 1.25, 3, 0, 2.5, 1, 2, Citytexture, 1, 1
    CreatePlaneInScene -5, 1.25, 3, 0, 2.5, 1, 2, Citytexture, 1, 1
    CreatePlaneInScene 5, 1.25, -3, 0, 2.5, 1, 2, Citytexture, 1, 1
    CreatePlaneInScene -5, 1.25, -3, 0, 2.5, 1, 2, Citytexture, 1, 1
    
    carlooked = car(1)
    carlookedpos.x = carpos(1).x
    carlookedpos.y = carpos(1).y
    carlookedpos.Z = carpos(1).Z

End Sub
Private Sub LoadTexture(surf As Direct3DRMTexture3, file As String, w As Long, h As Long)
    Set surf = RMC.D3DRM.CreateTextureFromSurface(RMC.CreateDDSFromCDIBFILE4(file, w, h, USESYSTEMMEMORY))
    surf.SetName UCase(file)
    surf.GenerateMIPMap
    surf.SetCacheOptions 0, D3DRMTEXTURE_STATIC
End Sub
Private Sub inittexture()
    Dim q As Single 'quality
    q = 8#
    
    Call LoadTexture(BuildTexture(0), "build1.gif", 32 * q, 32 * q)
    Call LoadTexture(BuildTexture(1), "build2.gif", 32 * q, 32 * q)
    Call LoadTexture(BuildTexture(2), "build3.gif", 32 * q, 32 * q)
    Call LoadTexture(BuildTexture(3), "build4.gif", 32 * q, 32 * q)
    Call LoadTexture(BuildTexture(4), "build5.jpg", 32 * q, 32 * q)
    Call LoadTexture(BuildTexture(5), "build6.jpg", 32 * q, 32 * q)
    Call LoadTexture(GroundTexture, "floor1.gif", 64 * q, 64 * q)
    Call LoadTexture(FloorTexture, "floor0.jpg", 16 * q, 16 * q)
    Call LoadTexture(Septexture, "sep2.jpg", 32 * q, 32 * q)
    Call LoadTexture(Citytexture, "City.gif", 16 * q, 16 * q)
    Call LoadTexture(WP(0), "mi.jpg", 32 * q, 32 * q)
    Call LoadTexture(WP(1), "c.jpg", 32 * q, 32 * q)
    Call LoadTexture(WP(2), "ti.jpg", 32 * q, 32 * q)
    Call LoadTexture(WP(3), "w.jpg", 32 * q, 32 * q)
    Call LoadTexture(RoofTexture, "roof.bmp", 32 * q, 32 * q)

End Sub
Private Sub Pic_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = vbKeyEscape Then
        Unload Me
    End If
    Select Case keyCode
        Case vbKeyF1
            ViewMode = 0
        Case vbKeyF2
            ViewMode = 1
        Case vbKeyF3
            ViewMode = 2
        Case vbKeyF4
            ViewMode = 3
        Case vbKeyF5
            ViewMode = 4
        Case vbKeyF6
            ViewMode = 5
        Case vbKeyF7
            Call RMC.CameraFrame.AddRotation(D3DRMCOMBINE_REPLACE, 1, 1, 1, 0)
            RMC.CameraFrame.SetPosition Nothing, 1.1, 0.5, -80
            ViewMode = 6
         Case vbKeyF12
            ViewMode = 11
         Case vbKeyM
            If music = True Then music = False Else music = True
    End Select
End Sub

Private Sub RM_E_DirecXNotInstalled()
    MsgBox "DirectX7 is not installed", vbCritical
    End
End Sub
Private Sub RM_E_Error4(Errstr As String)
    MsgBox Errstr, vbCritical, "Error"
    End
End Sub
Private Sub RM_E_PostRender()
    On Local Error Resume Next
      RMC.BackBuffer.SetForeColor RGB(255, 255, 255)  'vbWhite
'      RMC.BackBuffer.DrawText 10, 10, "Frame/Sec:" + CStr(RMC.FPS), False
End Sub

'Creating World
Sub Createsep(x As Single, Z As Single)
    CreateboxInScene 2.5 + x, 0.25, 2.5 + (25 / 2) + Z, 0.5, 0.5, 25, Septexture, 10, 1, False
    CreateboxInScene 2.5 + x, 0.25, -2.5 - (25 / 2) + Z, 0.5, 0.5, 25, Septexture, 10, 1, False
    CreateboxInScene 2.5 + (25 / 2) + x, 0.25, 2.5 + Z, 25, 0.5, 0.5, Septexture, 10, 1, False
    CreateboxInScene -2.5 - (25 / 2) - x, 0.25, 2.5 + Z, 25, 0.5, 0.5, Septexture, 10, 1, False
    CreateboxInScene -2.5 - x, 0.25, 2.5 + (25 / 2) + Z, 0.5, 0.5, 25, Septexture, 10, 1, False
    CreateboxInScene -2.5 - x, 0.25, -2.5 - (25 / 2) + Z, 0.5, 0.5, 25, Septexture, 10, 1, False
    CreateboxInScene 2.5 + (25 / 2) + x, 0.25, -2.5 + Z, 25, 0.5, 0.5, Septexture, 10, 1, False
    CreateboxInScene -2.5 - (25 / 2) - x, 0.25, -2.5 + Z, 25, 0.5, 0.5, Septexture, 10, 1, False
End Sub
Sub CreateBuilding(x As Single, Z As Single, TID As Integer, w As Single, h As Single, D As Single, Optional TU As Single, Optional TV As Single)

    Dim c As Single
    h = h * 1.2
    w = w * 2
    D = D * 2
    c = 6
    If TU = 0 Then TU = 2
    If TV = 0 Then TV = 2
    
    CreateboxInScene c + x, h / 2, c + Z, w, h, D, BuildTexture(TID), TU, TV, True
    CreateboxInScene -c + x, h / 2, c + Z, w, h, D, BuildTexture(TID), TU, TV, True
    CreateboxInScene c + x, h / 2, -c + Z, w, h, D, BuildTexture(TID), TU, TV, True
    CreateboxInScene -c + x, h / 2, -c + Z, w, h, D, BuildTexture(TID), TU, TV, True
    
End Sub
Sub CreateboxInScene(ByVal PX As Single, ByVal PY As Single, ByVal PZ As Single, ByVal w As Single, ByVal h As Single, ByVal D As Single, ByRef Texture As Direct3DRMTexture3, ByVal TU As Single, ByVal TV As Single, Optional insidep As Boolean)
    
    w = w / 2
    h = h / 2
    D = D / 2
    Dim mout As Direct3DRMMeshBuilder3
    Dim min As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Dim frm As Direct3DRMFrame3
    Dim OUTT As Direct3DRMTexture3 'outlook texture
    Dim INTT As Direct3DRMTexture3 'in texture
    
    Set frm = RMC.D3DRM.CreateFrame(Nothing)
    Set OUTT = Texture
    Set INTT = Texture
    
    Set mout = RMC.D3DRM.CreateMeshBuilder
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex -w, -h, D
        .AddVertex w, -h, D
        .AddVertex w, h, D
        .AddVertex -w, h, D
    End With
    With f
        .SetTextureCoordinates 0, TU, TV
        .SetTextureCoordinates 1, 0, TV
        .SetTextureCoordinates 2, 0, 0
        .SetTextureCoordinates 3, TU, 0
    End With
    f.SetTexture OUTT
    mout.AddFace f
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex -w, h, -D
        .AddVertex w, h, -D
        .AddVertex w, -h, -D
        .AddVertex -w, -h, -D
    End With
    With f
        .SetTextureCoordinates 0, 0, 0
        .SetTextureCoordinates 1, TU, 0
        .SetTextureCoordinates 2, TU, TV
        .SetTextureCoordinates 3, 0, TV
    End With
    f.SetTexture OUTT
    mout.AddFace f
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex w, -h, -D
        .AddVertex w, h, -D
        .AddVertex w, h, D
        .AddVertex w, -h, D
    End With
    With f
        .SetTextureCoordinates 0, 0, TV
        .SetTextureCoordinates 1, 0, 0
        .SetTextureCoordinates 2, TU, 0
        .SetTextureCoordinates 3, TU, TV
    End With
    f.SetTexture OUTT
    mout.AddFace f
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex -w, -h, D
        .AddVertex -w, h, D
        .AddVertex -w, h, -D
        .AddVertex -w, -h, -D
    End With
    With f
        .SetTextureCoordinates 0, 0, TV
        .SetTextureCoordinates 1, 0, 0
        .SetTextureCoordinates 2, TU, 0
        .SetTextureCoordinates 3, TU, TV
    End With
    f.SetTexture OUTT
    mout.AddFace f
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex -w, h, D
        .AddVertex w, h, D
        .AddVertex w, h, -D
        .AddVertex -w, h, -D
    End With
    With f
        .SetTextureCoordinates 0, 0, 0
        .SetTextureCoordinates 1, TU, 0
        .SetTextureCoordinates 2, TU, TV
        .SetTextureCoordinates 3, 0, TV
    End With
    f.SetTexture RoofTexture
    mout.AddFace f
    Set f = RMC.D3DRM.CreateFace
    With f
        .AddVertex -w, -h, -D
        .AddVertex w, -h, -D
        .AddVertex w, -h, D
        .AddVertex -w, -h, D
    End With
    With f
        .SetTextureCoordinates 0, 0, 0
        .SetTextureCoordinates 1, TU, 0
        .SetTextureCoordinates 2, TU, TV
        .SetTextureCoordinates 3, 0, TV
    End With
    f.SetTexture OUTT
    mout.AddFace f
    mout.SetPerspective D_TRUE
    mout.Optimize
    frm.AddVisual mout
    
    If insidep = True Then
        Set min = RMC.D3DRM.CreateMeshBuilder
        min.SetQuality D3DRMRENDER_UNLITFLAT
    'BACK
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex -w, h, D
            .AddVertex w, h, D
            .AddVertex w, -h, D
            .AddVertex -w, -h, D
        End With
        With f
            .SetTextureCoordinates 0, TU, 0
            .SetTextureCoordinates 1, 0, 0
            .SetTextureCoordinates 2, 0, TV
            .SetTextureCoordinates 3, TU, TV
        End With
        f.SetTexture INTT
        min.AddFace f
    'FRONT
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex -w, -h, -D
            .AddVertex w, -h, -D
            .AddVertex w, h, -D
            .AddVertex -w, h, -D
        End With
        With f
            .SetTextureCoordinates 0, 0, TV
            .SetTextureCoordinates 1, TU, TV
            .SetTextureCoordinates 2, TU, 0
            .SetTextureCoordinates 3, 0, 0
        End With
        f.SetTexture INTT
        min.AddFace f
    'left
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex w, -h, D
            .AddVertex w, h, D
            .AddVertex w, h, -D
            .AddVertex w, -h, -D
        End With
        With f
            .SetTextureCoordinates 0, TU, TV
            .SetTextureCoordinates 1, TU, 0
            .SetTextureCoordinates 2, 0, 0
            .SetTextureCoordinates 3, 0, TV
        End With
        f.SetTexture INTT
        min.AddFace f
    'right
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex -w, -h, -D
            .AddVertex -w, h, -D
            .AddVertex -w, h, D
            .AddVertex -w, -h, D
        End With
        With f
            .SetTextureCoordinates 0, TU, TV
            .SetTextureCoordinates 1, TU, 0
            .SetTextureCoordinates 2, 0, 0
            .SetTextureCoordinates 3, 0, TV
        End With
        f.SetTexture INTT
        min.AddFace f
    'Top
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex -w, h, -D
            .AddVertex w, h, -D
            .AddVertex w, h, D
            .AddVertex -w, h, D
        End With
        With f
            .SetTextureCoordinates 0, 0, TV
            .SetTextureCoordinates 1, TU, TV
            .SetTextureCoordinates 2, TU, 0
            .SetTextureCoordinates 3, 0, 0
        End With
        f.SetTexture INTT
        min.AddFace f
    'Bottom
        Set f = RMC.D3DRM.CreateFace
        With f
            .AddVertex -w, -h, D
            .AddVertex w, -h, D
            .AddVertex w, -h, -D
            .AddVertex -w, -h, -D
        End With
        With f
            .SetTextureCoordinates 0, 0, TV
            .SetTextureCoordinates 1, TU, TV
            .SetTextureCoordinates 2, TU, 0
            .SetTextureCoordinates 3, 0, 0
        End With
        f.SetTexture INTT
        min.AddFace f
        
        min.SetPerspective D_TRUE
        min.Optimize
        Dim fin As Direct3DRMFrame3
        Set fin = RMC.D3DRM.CreateFrame(frm)
        Dim mo As D3DRMMATERIALOVERRIDE
        With mo
            .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
            .dcDiffuse.a = 0.6
        End With
        fin.SetMaterialOverride mo
        fin.AddVisual min
    End If
    frm.SetPosition Nothing, PX, PY, PZ
    RMC.SceneFrame.AddVisual frm
End Sub
Sub CreatePlaneInScene(PX As Single, PY As Single, PZ As Single, Width As Single, Height As Single, depth As Single, nSides As Integer, Texture As Direct3DRMTexture3, TU As Single, TV As Single)
    
    Dim msh As Direct3DRMMeshBuilder3
    Dim frm As Direct3DRMFrame3
    
    If depth = 0 Then
        Set msh = RMC.CreateSheetMesh(nSides, Width, Height, TU, TV)
    ElseIf Width = 0 Then
        Set msh = RMC.CreateSheetMesh(nSides, Height, depth, TU, TV)
    ElseIf Height = 0 Then
        Set msh = RMC.CreateSheetMesh(nSides, depth, Width, TU, TV)
    End If
    Set frm = RMC.D3DRM.CreateFrame(Nothing)
    msh.SetMaterial mat
    msh.SetTexture Texture
    msh.SetPerspective D_TRUE
    msh.Optimize
    frm.AddVisual msh
    If depth = 0 Then
        frm.AddRotation D3DRMCOMBINE_REPLACE, 0, 0, 1, -3.14 / 2
    ElseIf Width = 0 Then
        frm.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, -3.14 / 2
    ElseIf Height = 0 Then
        frm.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, 3.14 / 2
    End If
    frm.SetPosition Nothing, PX, PY, PZ
    RMC.SceneFrame.AddVisual frm
    
    Set frm = Nothing
    Set msh = Nothing
End Sub
Sub createwallposter(PX As Single, PZ As Single, Texture As Direct3DRMTexture3)
    Dim s1 As Direct3DRMMeshBuilder3
    Dim s3 As Direct3DRMMeshBuilder3
    Dim frm As Direct3DRMFrame3
    Set s1 = RMC.CreateSheetMesh(2, 1, 0.5, 1, 1)
    s1.Translate PX, 0.5, PZ
    s1.SetTexture Septexture
    Set s3 = RMC.CreateSheetMesh(2, 3, 7, 1, 1)
    s3.Translate PX, 2.5, PZ
    s3.SetTexture Texture
    Set frm = RMC.D3DRM.CreateFrame(RMC.SceneFrame)
    frm.AddVisual s1
    frm.AddVisual s3
    
    If Abs(PX) < Abs(PZ) Then
        frm.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 3.142 / 2
    End If
End Sub
Sub createlightpot(x As Single, Z As Single, dirx, dirz, Texture As Direct3DRMTexture3)
    CreateboxInScene x, 1, Z, 0.1, 2, 0.1, Texture, 1, 1
    CreateboxInScene x, 1, Z, 0.1, 2, 0.1, Texture, 1, 1
    CreateboxInScene x, 1, Z, 0.1, 2, 0.1, Texture, 1, 1
    CreateboxInScene x, 1, Z, 0.1, 2, 0.1, Texture, 1, 1
    If dirz = 0 Then
       CreateboxInScene x + dirx, 2, Z + dirz, 0.4, 0.1, 0.2, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.4, 0.1, 0.2, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.4, 0.1, 0.2, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.4, 0.1, 0.2, Texture, 1, 1
    End If
    If dirx = 0 Then
       CreateboxInScene x + dirx, 2, Z + dirz, 0.2, 0.1, 0.4, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.2, 0.1, 0.4, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.2, 0.1, 0.4, Texture, 1, 1
       CreateboxInScene x + dirx, 2, Z + dirz, 0.2, 0.1, 0.4, Texture, 1, 1
    End If

End Sub
Sub createpotlight(x As Single, Z As Single, dirx, dirz)
    Set I_oD3MeshLamps1 = RMC.D3DRM.CreateMeshBuilder
    I_oD3MeshLamps1.LoadFromFile "l1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    I_oD3MeshLamps1.ScaleMesh 0.3, 0.1, 0.5
    I_oD3MeshLamps1.Optimize
    Set I_oD3Frame = RMC.D3DRM.CreateFrame(Nothing)
    If dirz = 0 Then
       I_oD3Frame.SetPosition Nothing, x, 2.22, Z - dirz
       I_oD3Frame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -180 * (pi / 180)
       If dirx > 0 Then
          I_oD3Frame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
       Else
          I_oD3Frame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -90 * (pi / 180)
       End If
       Set I_oD3Light = RMC.D3DRM.CreateLightRGB(D3DRMLIGHT_SPOT, 1, 1, 1)
       With I_oD3Light
            .SetLinearAttenuation 0.1
            .SetRange 2
            .SetUmbra 0.6
            .SetPenumbra 0.8
       End With
       I_oD3Frame.AddLight I_oD3Light
       I_oD3Frame.AddVisual I_oD3MeshLamps1
    End If
    If dirx = 0 Then
       I_oD3Frame.SetPosition Nothing, x - dirx, 2.22, Z
       I_oD3Frame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -180 * (pi / 180)
    End If
    RMC.SceneFrame.AddVisual I_oD3Frame
End Sub
