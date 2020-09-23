Attribute VB_Name = "modMain"
'//------------------------------------------------------------------------------
'// Name:modMain
'// Desc:the module that starts the application and ends the application
'//------------------------------------------------------------------------------

'// Doesn't allow me to use variables without defining them first
Option Explicit

'//------------------------------------------------------------------------------
'// Global API's
'//------------------------------------------------------------------------------

'// Sets the minimum timer resolution for the application
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) _
As Long

'// Resets the minimum timer resolution for the application to default
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) _
As Long

'// Returns the system time, in milliseconds
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'// Allows the app. to user input from the mouse or keyboard
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) _
As Long

'// Hides and shows the mouse icon
Public Declare Function ShowCursor Lib "user32" (ByVal Show As Boolean) As Long

'//------------------------------------------------------------------------------
'// Global User-Defined Types
'//------------------------------------------------------------------------------

'// A udt that holds the x and y coords of my ship's parts
Public Type ShipCoordsType
  dX As Double
  dY As Double
End Type

'// A udt that holds the different parts of my ship including direction
'// inertia and gravity
Public Type ShipPartsType
  udtTop          As ShipCoordsType
  udtCenter       As ShipCoordsType
  udtLeft         As ShipCoordsType
  udtRight        As ShipCoordsType
  udtBottom       As ShipCoordsType
  dRad            As Double
  dVelocityY      As Double
  dVelocityRotate As Double
  dVelocityX      As Double
End Type

'//------------------------------------------------------------------------------
'// Global Constants
'//------------------------------------------------------------------------------

'// My constant for PI
Const g_dPI As Double = 3.14159265358979

'// My constant to convert degrees to radians, this comes from PI/180
Const g_dDEG_TO_RAD As Double = 1.74532925199433E-02

'// My constant to convert radians to degrees, this comes from 1/(PI/180)
Const g_dRAD_TO_DEG As Double = 57.2957795130823

'// The constant for the diameter that the ship rotates about
Const g_lDIAMETER As Long = 12

'// The constnat that would make my ship equlateral
Const g_dEQULATERAL As Double = g_lDIAMETER * 0.5 * 1.73205080756888

'// The constant that I increase/decrease the velocity of the rear booster by
Const g_dVEL_BOOST As Double = 0.05

'// The constant that I increase/decrease the velocity of the front booster by
Const g_dVEL_REV_BOOST As Double = 0.0125

'// The constant for the maximum velocity of the rear booster
Const g_dMAX_VEL_BOOST As Double = 100

'// The constant that I increase/decrease the velocity of the turning boosters by
Const g_dVEL_TURN As Double = g_dDEG_TO_RAD * 0.1

'// The constant rate that gravity pulls the ship down by
Const g_dGRAVITY  As Double = 0.0125

'// The constant that holds the text for when my game is paused
Const g_sPAUSE As String = "PAUSED PRESS P OR SPACE TO UNPAUSE"

'// The constant that holds the text for when the game is on level 0
Const g_sHOW_TO_CONTINUE As String = "PRESS SPACE TO CONTINUE"

'// The games refresh rate in ms
Const g_lINTERVAL As Long = 10

'// The number of levels in the game + 2 which include the win and begin pics
Const g_lMAX_LEVEL As Long = 25

'// The XSpeed that will cause you to crash if you hit green
Const g_dCRASH_X As Double = 0.5

'// The YSpeed that will cause you to crash if you hit green
Const g_lCRASH_Y As Long = 1

'// The Angle that will cause you to crash if you hit green with your left side
Const g_dCRASH_LEFT_ANGLE As Double = 247.5

'// The Angle that will cause you to crash if you hit green with your right side
Const g_dCRASH_RIGHT_ANGLE As Double = 292.5

'//------------------------------------------------------------------------------
'// Global Variables
'//------------------------------------------------------------------------------

'// Used to keep the data that is from DirectX7
Public g_dx7 As DirectX7

'// Used to keep the data that is from DirectDraw7
Public g_dd7 As DirectDraw7

'// Used to keep the ShipPartsType data
Public g_udtShip As ShipPartsType
 
'// Stops you from holding a button too long
Public g_bKeyHoldStop As Boolean

'// Used To Start The Game Loop
Public g_bInit As Boolean

'// Used to pause the game without ending the program
Public g_bPause As Boolean

'// Used to save the last tick of timegettime
Public g_lLastTick As Long

'// Used to keep track of which level it is
Public g_lLevel As Long

'// Used to enable level select after you beat the game
Public g_bLevelSelect As Boolean

'//------------------------------------------------------------------------------
'// Name: ClearKeys
'// Desc: Used to clear the keys used by GetAsyncKeyState in order to prevent
'//       misfire at beginning of program
'//------------------------------------------------------------------------------
Private Sub ClearKeys()
  Dim lKey As Long
  For lKey = 0& To 255&
    GetAsyncKeyState lKey
  Next lKey
End Sub

'//------------------------------------------------------------------------------
'// Name: MoveShip
'// Desc: Moves the center of the ship
'//------------------------------------------------------------------------------
Private Sub MoveShip()
  With g_udtShip
    
    '// Tells the game that if it reads the arrow keys being pressed then to
    '// have the ship react accordingly
    If GetAsyncKeyState(vbKeyUp) Then
      .dVelocityX = .dVelocityX + Cos(.dRad) * g_dVEL_BOOST
      .dVelocityY = .dVelocityY + Sin(.dRad) * g_dVEL_BOOST
    End If
    If GetAsyncKeyState(vbKeyDown) Then
      .dVelocityX = .dVelocityX - Cos(.dRad) * g_dVEL_REV_BOOST
      .dVelocityY = .dVelocityY - Sin(.dRad) * g_dVEL_REV_BOOST
    End If
    If GetAsyncKeyState(vbKeyLeft) Then
      .dVelocityRotate = .dVelocityRotate - g_dVEL_TURN
    ElseIf .dVelocityRotate < 0& Then
      .dVelocityRotate = .dVelocityRotate + g_dVEL_TURN
      If .dVelocityRotate > 0& Then
        .dVelocityRotate = 0&
      End If
    End If
    
    If GetAsyncKeyState(vbKeyRight) Then
      .dVelocityRotate = .dVelocityRotate + g_dVEL_TURN
    ElseIf .dVelocityRotate > 0& Then
      .dVelocityRotate = .dVelocityRotate - g_dVEL_TURN
      If .dVelocityRotate < 0& Then
        .dVelocityRotate = 0&
      End If
    End If
    
    '// Creates gravity for the ship to fight against
    If g_lLevel <> g_lMAX_LEVEL Then
      .dVelocityY = .dVelocityY + g_dGRAVITY
    End If
    
    '// Sets the max angle to 360 degrees and the min angle to 1
    If .dRad > 360& * g_dDEG_TO_RAD Then
      .dRad = 1& * g_dDEG_TO_RAD
    End If
    
    If .dRad < 1& * g_dDEG_TO_RAD Then
      .dRad = 360& * g_dDEG_TO_RAD
    End If
    
    
    '// Sets a Maximum Velocity for X, Y, and Rotate
    If .dVelocityX <= -g_dMAX_VEL_BOOST Then
      .dVelocityX = -g_dMAX_VEL_BOOST
    End If
    If .dVelocityX >= g_dMAX_VEL_BOOST Then
      .dVelocityX = g_dMAX_VEL_BOOST
    End If
    If .dVelocityY <= -g_dMAX_VEL_BOOST Then
      .dVelocityY = -g_dMAX_VEL_BOOST
    End If
    If .dVelocityY >= g_dMAX_VEL_BOOST Then
      .dVelocityY = g_dMAX_VEL_BOOST
    End If
    If .dVelocityRotate <= -g_dPI Then
      .dVelocityRotate = -g_dPI
    End If
    If .dVelocityRotate >= g_dPI Then
      .dVelocityRotate = g_dPI
    End If
    
    '// Makes the ship always react to the current velocity changes, giving the
    '// effect of gravity and inertia
    .udtCenter.dX = .udtCenter.dX + .dVelocityX
    .udtCenter.dY = .udtCenter.dY + .dVelocityY
    .dRad = .dRad + .dVelocityRotate
  End With
End Sub


'//------------------------------------------------------------------------------
'// Name: BuildShip
'// Desc: Plots the coordinates of my lander around the center of the ship
'//------------------------------------------------------------------------------
Private Sub BuildShip()
  With g_udtShip
    
    '// Sets the coords of the top and bottom of my ship around the center of
    '// my ship
    .udtTop.dX = .udtCenter.dX + g_lDIAMETER * Cos(.dRad)
    .udtTop.dY = .udtCenter.dY + g_lDIAMETER * Sin(.dRad)
    
    .udtBottom.dX = .udtCenter.dX + g_lDIAMETER * CDbl(0.5) * Cos(.dRad + g_dPI)
    .udtBottom.dY = .udtCenter.dY + g_lDIAMETER * CDbl(0.5) * Sin(.dRad + g_dPI)
    
    '// Sets the coords of the left and right of my ship around the bottom of
    '// my ship
    .udtLeft.dX = .udtBottom.dX + g_dEQULATERAL * CDbl(0.8) * _
                  Cos(.dRad - g_dPI * 0.5)
    .udtLeft.dY = .udtBottom.dY + g_dEQULATERAL * CDbl(0.8) * _
                  Sin(.dRad - g_dPI * 0.5)
    
    .udtRight.dX = .udtBottom.dX + g_dEQULATERAL * CDbl(0.8) * _
                   Cos(.dRad + g_dPI * 0.5)
    .udtRight.dY = .udtBottom.dY + g_dEQULATERAL * CDbl(0.8) * _
                   Sin(.dRad + g_dPI * 0.5)
  
  End With
End Sub


'//------------------------------------------------------------------------------
'// Name: Loser
'// Desc: Called if you crash
'//------------------------------------------------------------------------------
Private Sub Loser()
  g_bPause = True
  LoadLevel
End Sub


'//------------------------------------------------------------------------------
'// Name: Winner
'// Desc: Called if you beat a level
'//------------------------------------------------------------------------------
Private Sub Winner()
  g_bPause = True
  g_lLevel = g_lLevel + 1&
  LoadLevel
End Sub



'//------------------------------------------------------------------------------
'// Name: LoadLevel
'// Desc: Loads a bitmap from "Mad Lander.RES" to represent a new level
'//------------------------------------------------------------------------------
Private Sub LoadLevel()
  '// Resets the ship
  With g_udtShip
       .udtCenter.dX = 40&
       .udtCenter.dY = 20&
         .dVelocityX = 0&
         .dVelocityY = 0&
    .dVelocityRotate = 0&
    .dRad = 270& * g_dDEG_TO_RAD
  End With
  BuildShip
  
  '// Sets the level
  Select Case g_lLevel
    Case Is = -1&
      g_lLevel = g_lMAX_LEVEL
      frmScreen.Picture = LoadResPicture("WINNER", vbResBitmap)
    Case 0&
      ElapsedTime
      frmScreen.Picture = LoadResPicture("HOW_TO_PLAY", vbResBitmap)
    Case 1& To g_lMAX_LEVEL - 1&
      frmScreen.Picture = LoadResPicture("LEVEL" & g_lLevel, vbResBitmap)
    Case g_lMAX_LEVEL
      frmScreen.Picture = LoadResPicture("WINNER", vbResBitmap)
      '// enables the levelselect sub
      g_bLevelSelect = True
    Case Is = g_lMAX_LEVEL + 1&
      g_lLevel = 0&
      frmScreen.Picture = LoadResPicture("HOW_TO_PLAY", vbResBitmap)
  End Select

End Sub

'//------------------------------------------------------------------------------
'// Name: Detection
'// Desc: Watches to see if the ship is hiting one of the sides of the level
'//       with one of it's tips
'//------------------------------------------------------------------------------
Private Sub Detection()
  With g_udtShip
    '// Checks to see if you hit the blue walls, if you do you lose
    If frmScreen.Point(.udtTop.dX, .udtTop.dY) = vbBlue Then
      Loser
    End If
    If frmScreen.Point(.udtLeft.dX, .udtLeft.dY) = vbBlue Then
      Loser
    End If
    If frmScreen.Point(.udtRight.dX, .udtRight.dY) = vbBlue Then
      Loser
    End If
    
    '// If you hit the green wall with the top of your ship you lose
    If frmScreen.Point(.udtTop.dX, .udtTop.dY) = vbGreen Then
      Loser
    End If
    
    '// If you hit the green wall with the right amount of speed and at the
    '// right angle then you can go to the next level
    If frmScreen.Point(.udtLeft.dX, .udtLeft.dY) = vbGreen Then
      If .dRad * g_dRAD_TO_DEG > g_dCRASH_LEFT_ANGLE And _
         .dVelocityY < g_lCRASH_Y And _
         .dVelocityX < g_dCRASH_X And .dVelocityX > -g_dCRASH_X Then
        Winner
      Else
        Loser
      End If
    End If
    If frmScreen.Point(.udtRight.dX, .udtRight.dY) = vbGreen Then
      If .dRad * g_dRAD_TO_DEG < g_dCRASH_RIGHT_ANGLE And _
         .dVelocityY < g_lCRASH_Y And _
         .dVelocityX < g_dCRASH_X And .dVelocityX > -g_dCRASH_X Then
        Winner
      Else
        Loser
      End If
    End If

    '// if you fly off one side of the screen then you will come back around the
    '// other side
    If .udtCenter.dX < -g_lDIAMETER Then
      .udtCenter.dX = frmScreen.ScaleWidth + g_lDIAMETER
      BuildShip
    End If
    If .udtCenter.dX > frmScreen.ScaleWidth + g_lDIAMETER Then
      .udtCenter.dX = -g_lDIAMETER
      BuildShip
    End If
    
    If .udtCenter.dY < -g_lDIAMETER Then
      .udtCenter.dY = frmScreen.ScaleHeight + g_lDIAMETER
      BuildShip
    End If
    If .udtCenter.dY > frmScreen.ScaleHeight + g_lDIAMETER Then
      .udtCenter.dY = -g_lDIAMETER
      BuildShip
    End If
  End With
End Sub



'//------------------------------------------------------------------------------
'// Name: Draw
'// Desc: Draws my lines and my text to the screen when called
'//------------------------------------------------------------------------------
Private Sub Draw()
  '// Writes if the game is paused with white text
  frmScreen.ForeColor = vbWhite
  
  If g_bPause = True Then frmScreen.Print Space$(22) & g_sPAUSE & _
                                          "  ElapsedTime " & ElapsedTime
  
  With g_udtShip
    '// Gives useful info to the user
    '// X and Y speed is set to Pixels Per Millisecond
    frmScreen.ForeColor = vbWhite
    If g_bPause = False Then
      frmScreen.Print Space$(22&) & _
                      CStr("Level:") & g_lLevel & CStr("  ") & ElapsedTime & _
                      CStr("  Angle:") & CLng(.dRad * g_dRAD_TO_DEG) & _
                      CStr("  XSpeed:") & CLng(.dVelocityX * 10&) & _
                      CStr("  YSpeed:") & CLng(.dVelocityY * 10&)
    End If
    
    
    '// Draws the bottom lines of my ship with Green or random shades of red to
    '// Yellow if up is pressed to give booster effect
    If GetAsyncKeyState(vbKeyUp) Then
      frmScreen.ForeColor = RGB(255&, CLng(Rnd) * 255&, 0&)
    Else
      frmScreen.ForeColor = vbGreen
    End If
    frmScreen.Line (.udtCenter.dX, .udtCenter.dY)-(.udtRight.dX, .udtRight.dY)
    frmScreen.Line (.udtCenter.dX, .udtCenter.dY)-(.udtLeft.dX, .udtLeft.dY)
    
    '// draws the top of the ship with green
    frmScreen.ForeColor = vbGreen
    frmScreen.Line (.udtTop.dX, .udtTop.dY)-(.udtLeft.dX, .udtLeft.dY)
    frmScreen.Line (.udtTop.dX, .udtTop.dY)-(.udtRight.dX, .udtRight.dY)
    frmScreen.Line (.udtTop.dX, .udtTop.dY)-(.udtCenter.dX, .udtCenter.dY)
  End With
End Sub

'//------------------------------------------------------------------------------
'// Name: SetScreenRes
'// Desc: When called it uses DirectX7 to convert the screen res
'//------------------------------------------------------------------------------
Private Sub SetScreenRes(Frm As Form, Width As Long, Height As Long, _
                         BitsPerPixel As Long, RefreshRate)
  Set g_dx7 = New DirectX7
  Set g_dd7 = g_dx7.DirectDrawCreate(LenB("") = 0&)
  g_dd7.SetCooperativeLevel Frm.hWnd, DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE Or _
  DDSCL_FULLSCREEN
  g_dd7.SetDisplayMode Width, Height, BitsPerPixel, RefreshRate, DDSDM_DEFAULT
  Frm.Width = Width * 15&
  Frm.Height = Height * 15&
End Sub

'//------------------------------------------------------------------------------
'// Name: ElapsedTime
'// Desc: Returns the amount of time you have been playing this game
'//------------------------------------------------------------------------------
Private Function ElapsedTime() As String
  Static lasttick As Long
  Static h        As Long
  Static tm       As Long
  Static m        As Long
  Static ts       As Long
  Static s        As Long
  
  If g_lLevel = 0 Then
    s = 0: ts = 0: m = 0: tm = 0: h = 0
  Else
    If g_lLevel <> g_lMAX_LEVEL Then
        If g_bPause = False Then
          If h < 1000 Then
            If Abs(CLng(Timer) - lasttick) >= 1 Then
              lasttick = CLng(Timer)
              s = s + 1
              If s > 9 Then
                s = 0
                ts = ts + 1
              End If
              If ts > 5 Then
                ts = 0
                m = m + 1
              End If
              If m > 9 Then
                m = 0
                tm = tm + 1
              End If
              If tm > 5 Then
                tm = 0
                h = h + 1
              End If
            End If
            ElapsedTime = h & ":" & tm & m & "." & ts & s
          Else
            ElapsedTime = "999:99.99"
          End If
        Else
          ElapsedTime = h & ":" & tm & m & "." & ts & s
        End If
      Else
      ElapsedTime = h & ":" & tm & m & "." & ts & s
    End If
  End If
End Function

'//------------------------------------------------------------------------------
'// Name: LevelSelect
'// Desc: To enable the use of the PageUp and PageDown keys to select a level
'//------------------------------------------------------------------------------

Private Sub LevelSelect()
  If g_bKeyHoldStop = False And GetAsyncKeyState(vbKeyPageUp) Then
    g_lLevel = g_lLevel + 1&
    Loser
  End If
  
  If g_bKeyHoldStop = False And GetAsyncKeyState(vbKeyPageDown) Then
    g_lLevel = g_lLevel - 1&
    Loser
  End If
End Sub




'//------------------------------------------------------------------------------
'// Name: Main
'// Desc: The starting/ending point of the application
'//------------------------------------------------------------------------------
Private Sub Main()
  '// Loads every thing that needs to be loaded in the beginning of the program
  timeBeginPeriod 1&
  frmScreen.Show
  ShowCursor False
  ClearKeys
  g_bInit = True
  g_udtShip.dRad = 270 * g_dDEG_TO_RAD
  frmScreen.DrawWidth = 3&
  g_bPause = False
  frmScreen.BackColor = vbBlack
  Randomize CDbl(Now) + CDbl(Timer)
  BuildShip
  LoadLevel
  SetScreenRes frmScreen, 640&, 480&, 16&, 0&
  
  '// Sets up a loop for the game that keeps a interval of 10ms for most
  '// computers
  Do
    DoEvents
    If Abs(timeGetTime - g_lLastTick) >= g_lINTERVAL Then
      g_lLastTick = timeGetTime
      
      '// This allows you to move on from the howtoplay level by pressing space
      If g_lLevel = 0 And GetAsyncKeyState(vbKeySpace) Then
        g_lLevel = g_lLevel + 1
        LoadLevel
      End If
      
      '// toggles pause/unpause when you press the "P"/"SPACE" key
      If g_bKeyHoldStop = False And (GetAsyncKeyState(vbKeyP) _
         Or GetAsyncKeyState(vbKeySpace)) Then
        g_bPause = Not g_bPause
      End If
      
      '// Once level select is enabled then call the LevelSelect Sub
      If g_bLevelSelect = True Then
        LevelSelect
      End If
      
      
      '// Stops you from holding a button too long
      If GetAsyncKeyState(vbKeyP) Or GetAsyncKeyState(vbKeySpace) _
         Or GetAsyncKeyState(vbKeyPageUp) Or GetAsyncKeyState(vbKeyPageDown) Then
        If g_bKeyHoldStop = False Then
          g_bKeyHoldStop = True
        End If
      Else
        g_bKeyHoldStop = False
      End If
      
            
      '// Stops the loop allowing the game to unload safely when you press the
      '// "ESC" key
      If GetAsyncKeyState(vbKeyEscape) Then
        g_bInit = False
      End If
           
      '// Runs the game while unpaused
      If g_bPause = False Then
        MoveShip
        BuildShip
      End If
      
      '// The Game Detects if you hit something first then draws your ship
      '// so that the detection doesn't get messed up by it's own colors
      frmScreen.Cls
      Detection

      If g_lLevel <> 0& Then
        Draw
      Else
        frmScreen.ForeColor = vbWhite
        frmScreen.Print g_sHOW_TO_CONTINUE
      End If
      frmScreen.Refresh
      
      

    End If
  Loop While g_bInit = True
  
  '// Sets the timer resolution back to normal and unloads the program
  ShowCursor True
  timeEndPeriod 1&
  Unload frmScreen
End Sub
