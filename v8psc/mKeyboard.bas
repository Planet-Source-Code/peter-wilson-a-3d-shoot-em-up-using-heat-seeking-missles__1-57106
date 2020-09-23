Attribute VB_Name = "mKeyboard"
Option Explicit

Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_CANCEL = &H3
Private Const VK_MBUTTON = &H4
Private Const VK_BACK = &H8
Private Const VK_TAB = &H9
Private Const VK_CLEAR = &HC
Private Const VK_RETURN = &HD
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12
Private Const VK_PAUSE = &H13
Private Const VK_CAPITAL = &H14
Private Const VK_ESCAPE = &H1B
Private Const VK_SPACE = &H20
Private Const VK_PRIOR = &H21
Private Const VK_NEXT = &H22
Private Const VK_END = &H23
Private Const VK_HOME = &H24
Private Const VK_LEFT = &H25
Private Const VK_UP = &H26
Private Const VK_RIGHT = &H27
Private Const VK_DOWN = &H28
Private Const VK_SELECT = &H29
Private Const VK_PRINT = &H2A
Private Const VK_EXECUTE = &H2B
Private Const VK_SNAPSHOT = &H2C
Private Const VK_INSERT = &H2D
Private Const VK_DELETE = &H2E
Private Const VK_HELP = &H2F
Private Const VK_0 = &H30
Private Const VK_1 = &H31
Private Const VK_2 = &H32
Private Const VK_3 = &H33
Private Const VK_4 = &H34
Private Const VK_5 = &H35
Private Const VK_6 = &H36
Private Const VK_7 = &H37
Private Const VK_8 = &H38
Private Const VK_9 = &H39
Private Const VK_A = &H41
Private Const VK_B = &H42
Private Const VK_C = &H43
Private Const VK_D = &H44
Private Const VK_E = &H45
Private Const VK_F = &H46
Private Const VK_G = &H47
Private Const VK_H = &H48
Private Const VK_I = &H49
Private Const VK_J = &H4A
Private Const VK_K = &H4B
Private Const VK_L = &H4C
Private Const VK_M = &H4D
Private Const VK_N = &H4E
Private Const VK_O = &H4F
Private Const VK_P = &H50
Private Const VK_Q = &H51
Private Const VK_R = &H52
Private Const VK_S = &H53
Private Const VK_T = &H54
Private Const VK_U = &H55
Private Const VK_V = &H56
Private Const VK_W = &H57
Private Const VK_X = &H58
Private Const VK_Y = &H59
Private Const VK_Z = &H5A
Private Const VK_STARTKEY = &H5B
Private Const VK_CONTEXTKEY = &H5D
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_MULTIPLY = &H6A
Private Const VK_ADD = &H6B
Private Const VK_SEPARATOR = &H6C
Private Const VK_SUBTRACT = &H6D
Private Const VK_DECIMAL = &H6E
Private Const VK_DIVIDE = &H6F
Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_NUMLOCK = &H90
Private Const VK_OEM_SCROLL = &H91
Private Const VK_OEM_1 = &HBA
Private Const VK_OEM_PLUS = &HBB
Private Const VK_OEM_COMMA = &HBC
Private Const VK_OEM_MINUS = &HBD
Private Const VK_OEM_PERIOD = &HBE
Private Const VK_OEM_2 = &HBF
Private Const VK_OEM_3 = &HC0
Private Const VK_OEM_4 = &HDB
Private Const VK_OEM_5 = &HDC
Private Const VK_OEM_6 = &HDD
Private Const VK_OEM_7 = &HDE
Private Const VK_OEM_8 = &HDF
Private Const VK_ICO_F17 = &HE0
Private Const VK_ICO_F18 = &HE1
Private Const VK_OEM102 = &HE2
Private Const VK_ICO_HELP = &HE3
Private Const VK_ICO_00 = &HE4
Private Const VK_ICO_CLEAR = &HE6
Private Const VK_OEM_RESET = &HE9
Private Const VK_OEM_JUMP = &HEA
Private Const VK_OEM_PA1 = &HEB
Private Const VK_OEM_PA2 = &HEC
Private Const VK_OEM_PA3 = &HED
Private Const VK_OEM_WSCTRL = &HEE
Private Const VK_OEM_CUSEL = &HEF
Private Const VK_OEM_ATTN = &HF0
Private Const VK_OEM_FINNISH = &HF1
Private Const VK_OEM_COPY = &HF2
Private Const VK_OEM_AUTO = &HF3
Private Const VK_OEM_ENLW = &HF4
Private Const VK_OEM_BACKTAB = &HF5
Private Const VK_ATTN = &HF6
Private Const VK_CRSEL = &HF7
Private Const VK_EXSEL = &HF8
Private Const VK_EREOF = &HF9
Private Const VK_PLAY = &HFA
Private Const VK_ZOOM = &HFB
Private Const VK_NONAME = &HFC
Private Const VK_PA1 = &HFD
Private Const VK_OEM_CLEAR = &HFE
Private Declare Function GetKeyboardState Lib "user32.dll" (lpKeyState As Byte) As Long

Public Sub GetKeyboardInput(strGameState As String, objPlayerOne As mdrPlayer)
    
    Dim keystates(0 To 255) As Byte
    Dim lngReturnValue As Long
    Dim intN As Integer
    
    Dim vectMoveForwardBackward As mdrVector4
    Dim vectMoveLeftRight As mdrVector4
    
    ' GetKeyboardState retrieves the state of every key on the keyboard and places
    ' the information into an array.
    lngReturnValue = GetKeyboardState(keystates(0))
    
    
    Select Case strGameState
        Case "demo_running" ' Program is currently running in the demo/credits state.
            
            ' Check for ESCAPE key during demo - ie. SkipDemo
            ' ================================================
            If (keystates(VK_ESCAPE) And 128) = 128 Then g_clsApplication.GameState = "demo_stop"


        Case "run_game" ' Program is currently running in the game/simulation state.
                        
            vectMoveForwardBackward = objPlayerOne.VPN
            vectMoveLeftRight = objPlayerOne.LeftRightVector
            
            If (keystates(VK_SHIFT) And 128) = 128 Then
                ' ===========================================================
                ' Multiply vectors for s-l-o-w precision movement (optional).
                ' ===========================================================
                vectMoveForwardBackward = VectorMultiplyByScalar(vectMoveForwardBackward, 1)
                vectMoveLeftRight = VectorMultiplyByScalar(vectMoveLeftRight, 1)
            ElseIf (keystates(VK_CONTROL) And 128) = 128 Then
                ' ================================================
                ' Multiply vectors for SPEED! Yippee!! (optional).
                ' ================================================
                vectMoveForwardBackward = VectorMultiplyByScalar(vectMoveForwardBackward, 32)
                vectMoveLeftRight = VectorMultiplyByScalar(vectMoveLeftRight, 32)
            Else
                vectMoveForwardBackward = VectorMultiplyByScalar(vectMoveForwardBackward, 8)
                vectMoveLeftRight = VectorMultiplyByScalar(vectMoveLeftRight, 8)
            End If
            
            ' ==============
            ' Move Forwards.
            ' ==============
            If (keystates(VK_W) And 128) = 128 Then
                objPlayerOne.WorldPosition.X = objPlayerOne.WorldPosition.X - vectMoveForwardBackward.X
                objPlayerOne.WorldPosition.Y = objPlayerOne.WorldPosition.Y - vectMoveForwardBackward.Y ' << Remove this line to restrict movement to 2D ground plane.
                objPlayerOne.WorldPosition.Z = objPlayerOne.WorldPosition.Z - vectMoveForwardBackward.Z
            End If
            
            ' ===============
            ' Move Backwards.
            ' ===============
            If (keystates(VK_S) And 128) = 128 Then
                objPlayerOne.WorldPosition.X = objPlayerOne.WorldPosition.X + vectMoveForwardBackward.X
                objPlayerOne.WorldPosition.Y = objPlayerOne.WorldPosition.Y + vectMoveForwardBackward.Y ' << Remove this line to restrict movement to 2D ground plane.
                objPlayerOne.WorldPosition.Z = objPlayerOne.WorldPosition.Z + vectMoveForwardBackward.Z
            End If
            
            ' ===================
            ' Straff to the Left.
            ' ===================
            If (keystates(VK_A) And 128) = 128 Then
                objPlayerOne.WorldPosition.X = objPlayerOne.WorldPosition.X - vectMoveLeftRight.X
                objPlayerOne.WorldPosition.Z = objPlayerOne.WorldPosition.Z - vectMoveLeftRight.Z
            End If
            
            ' ====================
            ' Straff to the Right.
            ' ====================
            If (keystates(VK_D) And 128) = 128 Then
                objPlayerOne.WorldPosition.X = objPlayerOne.WorldPosition.X + vectMoveLeftRight.X
                objPlayerOne.WorldPosition.Z = objPlayerOne.WorldPosition.Z + vectMoveLeftRight.Z
            End If
            
            ' ===================================================================
            ' Use spacebar to move upwards (I'll do jumping in the next version).
            ' ===================================================================
            If (keystates(VK_SPACE) And 128) = 128 Then
                objPlayerOne.WorldPosition.Y = objPlayerOne.WorldPosition.Y + 32
            End If
            
            
            ' Check for ESCAPE key during game play - ie. Go back to Main Menu / Settings.
            ' ============================================================================
            If (keystates(VK_ESCAPE) And 128) = 128 Then g_clsApplication.GameState = "quit"
            
                                    
            If objPlayerOne.WorldPosition.Y < 0 Then objPlayerOne.WorldPosition.Y = 0
            
    End Select
    

End Sub

