Attribute VB_Name = "modMouseOver"
    Option Explicit
    

    'Put this code in a module file
    'Put check procedure in a timer control

    'Sample
    'Private Sub Timer1_Timer()
    '
    '    On Error Resume Next
    '
    '    If IsMouseOver(Image1) Then
    '        do something
    '    Else
    '        do something else
    '    End If
    '
    'End Sub

        Type POINTAPI
            x                                                              As Long
            y                                                              As Long
        End Type

    Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

    Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

    Declare Function IsWindowVisible Lib "user32" (ByVal HWND As Long) As Long

    Declare Function ScreenToClient Lib "user32" (ByVal HWND As Long, lpPoint As POINTAPI) As Long

Public Function IsMouseOver(Ctl As Control) As Boolean
    
    Dim typPt            As POINTAPI
    Dim mOver            As Long
    Dim HWND             As Long
    Dim CtlLeft&
    Dim CtlTop&
    Dim CtlRight&
    Dim CtlBottom&
    
    On Error Resume Next
    
    'Initialize Variables
    HWND = 0
    Err.Number = 0
    
    'Get controls handle
    HWND = Ctl.HWND
    
    'if control does not have a handle, an error is raised

        If Err.Number > 0 Then
            'Get the handle of the control's parent control or form
            HWND = Ctl.Container.HWND
        
            'Get current cursor position
            Call GetCursorPos(typPt)
        
            'Get the handle of the control under these coordinates
            mOver = WindowFromPoint(typPt.x, typPt.y)
        
            'If the returned control handle is equal to the parent
            'control handle then the mouse is over that parent control

                If mOver <> HWND Then
                    IsMouseOver = False
                    Exit Function
                End If
        
            'Get the rect of the questioned control
            'If the window's scalemode property is Pixels
            'then remove the TwipsPerPixel calculations
            CtlLeft = Ctl.Left / Screen.TwipsPerPixelX
            CtlTop = Ctl.Top / Screen.TwipsPerPixelY
            CtlRight = (Ctl.Left + Ctl.Width) / Screen.TwipsPerPixelX
            CtlBottom = (Ctl.Top + Ctl.Height) / Screen.TwipsPerPixelY
        
            'Convert the mouse's screen position to the
            'mouse's parent control position
            Call ScreenToClient(HWND, typPt)
        
            'If the mouse is within the questioned control's
            'coordinates then the mouse is over the questioned control

                If typPt.y >= CtlTop And typPt.y <= CtlBottom And typPt.x >= CtlLeft And typPt.x <= CtlRight Then
                    IsMouseOver = True
                Else
                    IsMouseOver = False
                End If
        
            'Reset error number
            Err.Number = 0
        
            'Stop here
            Exit Function
        End If
    
    'Questioned control has a handle so check it directly
    
    'Reset Variables
    Err.Number = 0
    HWND = Ctl.HWND
    
    'Get current cursor position
    Call GetCursorPos(typPt)
    
    'Get the handle of the control under these coordinates
    mOver = WindowFromPoint(typPt.x, typPt.y)
    
    'If the returned control handle is equal to the questioned
    'control handle then the mouse is over that control

        If mOver = HWND Then
            IsMouseOver = True
        Else
            IsMouseOver = False
        End If
    
End Function
