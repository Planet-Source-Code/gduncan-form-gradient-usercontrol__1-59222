VERSION 5.00
Begin VB.UserControl Duncan_GradientBackground 
   BackColor       =   &H000000FF&
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   1590
   ToolboxBitmap   =   "Duncan_GradientBackground.ctx":0000
End
Attribute VB_Name = "Duncan_GradientBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What?
'Applies a gradient background to the parent form

'How?
'We create an image and load it into the .picture property of the form
'read on - it isnt laggy

'Why do it this way?
'I do it this way because I found drawing the gradient directly onto the forms DC
'would overwrite any Label controls on the form
'I also found that by putting my draw function in the Paint event of my form I got
'alot of background lag. If my app wasnt active and I moved another window in front of
'my app the repaint in my app would be called even though it wasnt necessary
'By applying an image rather than drawing on the DC I have removed that background lag,
'and made applying a gradient to a form as simple as dropping on this control.
'The only disadvantage this method had before was that it is processor intensive and
'caused lag on resizing. I have negated this by introducing buffers.
'The values of the buffers are set for what was good on my 2ghz pc

'What you need to know
'Given that creating a picture and drawing a large gradient can be processor intensive
'I have introduced what I call Size Buffers to help limit redrendering of the gradient image
'When a background is made it is made slightly larger than the form.
'Actually it is made at the size of the form + the buffer size %
'The buffer size is either high quality or low quality depending on what action is being performed
'When the user is sizing the form we use a big size buffer so that less frequent updates
'occur. And when they exit the size event we make the image again but this time with
'the high quality buffer.
'The buffer size represents how far the form needs to move before a rerender occurs
'meaning that a better representation of the gradient is applied
'Buffer sizes are set in the sBPhq and sBPlq constants
'If you do experience lag when sizing or high processor demands then make the
'buffer sizes larger. the larger the values the less times the background
'will be rendered.


'How to use?
'just drop this control anywhere on your form and it will render the background at run time
'only need one per form

'Who helped?
'Thanks to Paul Catton for his work on subclassing - sourced from planetsourcecode.com
'Thanks to the guys who wrote the gradient function - sourced from planetsourcecode.com

'When?
'Last Updated : Feb 2005

'======================================================================================================================================================
'MY DECLARES FOR THIS CONTROL
'======================================================================================================================================================
Private Const API_DIB_RGB_COLORS As Long = 0
Private Type tpBITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Declare Function API_StretchDIBits Lib "gdi32" Alias "StretchDIBits" _
        (ByVal hdc As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As tpBITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long
         
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private m_ColourTop As Long         'top colour of gradient
Private m_ColourBottom As Long      'bottom colour of gradient
Private m_R As RECT                 'size of the image we have created
'Buffer size represents how far the form needs to move before the background is remade
'This is a % of form size because large gradients take alot of processor power
'and so we want to render them less frequently
Private m_sBP As Single               'buffer size % being used
Private Const sBPhq As Single = 0.02   'buffer size %    High Quality
Private Const sBPlq As Single = 0.15   'buffer size %    Low Quality

Private m_Validated As Boolean      'should we subclass?

'======================================================================================================================================================
'SUBCLASSING DECLARES
'======================================================================================================================================================
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Messages
Private Const WM_NCPAINT As Long = &H85 'border changed
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_EXITSIZEMOVE = &H232



'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data

    Select Case uMsg
        'When entering a resize use the low quality buffering
        Case WM_ENTERSIZEMOVE
            m_sBP = sBPlq
            'Debug.Print "entering resize"
            
        'When exiting a resize use the high quality buffering
        Case WM_EXITSIZEMOVE
            m_sBP = sBPhq
            DoIt
            'Debug.Print "exiting resize"
            
        'When border sizes have changed check and see if a repaint is needed
        Case WM_NCPAINT
            DoIt
            'Debug.Print "NC painted " & Now
    End Select
End Sub
'======================================================================================================================================================
'Functions
'======================================================================================================================================================


'-----------------
'PUBLIC PROPERTIES
'-----------------
Public Property Get ColourTop() As OLE_COLOR
    ColourTop = m_ColourTop
End Property
Public Property Let ColourTop(lCol As OLE_COLOR)
    If lCol <> m_ColourTop Then
        m_ColourTop = lCol
        PropertyChanged "ColourTop"
        DrawTeaser
    End If
End Property

Public Property Get ColourBottom() As OLE_COLOR
    ColourBottom = m_ColourBottom
End Property
Public Property Let ColourBottom(lCol As OLE_COLOR)
    If lCol <> m_ColourBottom Then
        m_ColourBottom = lCol
        PropertyChanged "ColourBottom"
        DrawTeaser
    End If
End Property


'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Sub DoIt()
    'main routine
    If m_Validated Then
        'if we need a new image then generate it
        If NewBackgroundNeeded Then
            CreateBackground
        End If
    End If
End Sub

Private Function NewBackgroundNeeded() As Boolean
    'Is a new background image needed?
    Dim W As Long
    Dim H As Long
    Dim R As RECT
    
    'we dont want to have to recreate the background for a 1 pixel move
    'or for each pixel moved when a window is opened right up wide
    'that would overload the pc and lag the app
    'to get around this we set a buffer zone
    'what this means is that we make the background picture slightly larger
    'than we need in anticipation of a move, but not so large as to deminish
    'the gradient effect
    'On a form border size change event we compare the size of our image and
    'the new form size. if it is within tollerance we do nothing
    'if a new background is needed then we make it at form size + buffer dimendions

    GetClientRect UserControl.Parent.hwnd, R
    
    'is image too small compared to form?
    If m_R.Right < R.Right Or _
       m_R.Bottom < R.Bottom Then
       NewBackgroundNeeded = True
       'Debug.Print "NBN image is less than form " & m_R.Right & "<" & R.Right
       Exit Function
    End If
    
    'is image too large compared to form?
    'if it is we will lose resolution
    W = m_R.Right * (1 - (2 * m_sBP))
    H = m_R.Bottom * (1 - (2 * m_sBP))
    If R.Right < W Or _
       R.Bottom < H Then
       NewBackgroundNeeded = True
       'Debug.Print "NBN: image is too big for form " & R.Right & "<" & W
       Exit Function
    End If
    
End Function

Private Sub CreateBackground()
    'creates a bitmap
    'on that bitmap we draw the gradient
    'we then set that bitmap as the parents .picture property
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long
    Dim R As RECT
    Dim pic As StdPicture
        
    'INITIALISE VARS
    'get the window we are drawing on so we know what size we need to make picture
    GetClientRect UserControl.Parent.hwnd, R
    m_R.Top = -1
    m_R.Bottom = R.Bottom + (R.Bottom * m_sBP)
    m_R.Right = R.Right + (R.Right * m_sBP)
    'Debug.Print "Image Created. Screen=" & R.Right & "x" & R.Bottom & " BG=" & m_R.Right & "x" & m_R.Bottom
    'Debug.Print "BufferSize: W=" & m_R.Right - R.Right & " H=" & m_R.Bottom - R.Bottom
    
    'CREATE OBJECTS NEEDED
    'create a compatible device context
    hDCMemory = CreateCompatibleDC(UserControl.hdc)
    'create a compatible bitmap
    hBmp = CreateCompatibleBitmap(UserControl.hdc, m_R.Right, m_R.Bottom)
    If hBmp <> 0 Then
        'select the compatible bitmap into our compatible device context
        hBmpPrev = SelectObject(hDCMemory, hBmp)
        'draw our gradient onto the bitmap
        DrawTopDownGradient hDCMemory, m_R, ColourTop, ColourBottom
    End If
    
    'CLEAN UP
    'restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    'delete the memory DC
    DeleteDC hDCMemory

    'APPLY
    If hBmp <> 0 Then
        Set pic = CreateBitmapPicture(hBmp)
        If Not pic Is Nothing Then
            'dont apply unless we have something
            Set UserControl.Parent.Picture = pic
        End If
    End If
End Sub

Private Sub DrawTopDownGradient(hdc As Long, rc As RECT, ByVal lRGBColorFrom As Long, ByVal lRGBColorTo As Long)
'this code sourced from
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57192&lngWId=1
'this code is update v3
    Dim uBIH            As tpBITMAPINFOHEADER
    Dim lBits()         As Long
    Dim lColor          As Long
    
    Dim x               As Long
    Dim y               As Long
    Dim xEnd            As Long
    Dim yEnd            As Long
    Dim ScanlineWidth   As Long
    Dim yOffset         As Long
    
    Dim R               As Long
    Dim G               As Long
    Dim B               As Long
    Dim end_R           As Long
    Dim end_G           As Long
    Dim end_B           As Long
    Dim dR              As Long
    Dim dG              As Long
    Dim dB              As Long
    
    
    ' Split a RGB long value into components - FROM gradient color
    lRGBColorFrom = lRGBColorFrom And &HFFFFFF                      ' "SplitRGB"  by www.Abstractvb.com
    R = lRGBColorFrom Mod &H100&                                    ' Should be the fastest way in pur VB
    lRGBColorFrom = lRGBColorFrom \ &H100&                          ' See test on VBSpeed (http://www.xbeat.net/vbspeed/)
    G = lRGBColorFrom Mod &H100&                                    ' Btw: API solution with RTLMoveMem is slower ... ;)
    lRGBColorFrom = lRGBColorFrom \ &H100&
    B = lRGBColorFrom Mod &H100&
    
    ' Split a RGB long value into components - TO gradient color
    lRGBColorTo = lRGBColorTo And &HFFFFFF
    end_R = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_G = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_B = lRGBColorTo Mod &H100&
    
    
    '-- Loops bounds
    xEnd = rc.Right - rc.Left
    yEnd = rc.Bottom - rc.Top
    
    ' Check:  Top lower than Bottom ?
    If yEnd < 1 Then
    
        Exit Sub
    End If
    
    '-- Scanline width
    ScanlineWidth = xEnd + 1
    yOffset = -ScanlineWidth
    
    '-- Initialize array size
    ReDim lBits((xEnd + 1) * (yEnd + 1) - 1) As Long
       
    '-- Get color distances
    dR = end_R - R
    dG = end_G - G
    dB = end_B - B
       
    '-- Gradient loop over rectangle
    For y = 0 To yEnd
        
        '-- Calculate color and *y* offset
        lColor = B + (dB * y) \ yEnd + 256 * (G + (dG * y) \ yEnd) + 65536 * (R + (dR * y) \ yEnd)
        
        yOffset = yOffset + ScanlineWidth
        
        '-- *Fill* line
        For x = yOffset To xEnd + yOffset
            lBits(x) = lColor
        Next x
        
    Next y
    
    '-- Prepare bitmap info structure
    With uBIH
        .biSize = Len(uBIH)
        .biBitCount = 32
        .biPlanes = 1
        .biWidth = xEnd + 1
        .biHeight = -yEnd + 1
    End With
    
    '-- Finaly, paint *bits* onto given DC
    API_StretchDIBits hdc, _
            rc.Left, rc.Top, _
            xEnd, yEnd, _
            0, 0, _
            xEnd, yEnd, _
            lBits(0), _
            uBIH, _
            API_DIB_RGB_COLORS, _
            vbSrcCopy
            
End Sub

Private Function CreateBitmapPicture(ByVal hBmp As Long) As Picture
    Dim retval As Long
    Dim pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic)        ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp            ' Handle to bitmap
        .hPal = 0               ' Handle to palette (may be null)
    End With

    'Create the picture
    retval = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function

Private Sub DrawTeaser()
    'this colours the usercontrol so that the developer can see
    'what colours are selected at a glance
    Dim R As RECT
    
    'only do this if we are in usermode (design mode)
    If Not Ambient.UserMode Then
        If m_Validated Then
            GetClientRect UserControl.hwnd, R
            R.Top = R.Top - 1
            DrawTopDownGradient UserControl.hdc, R, ColourTop, ColourBottom
        Else
            'leave it red so they know its not working
        End If
    End If
    
End Sub

Private Function ValidateControl() As Boolean
    'For the control to work it needs two things
    '1) to confirm that usercontrol.parent.picture is valid
    '2) that this control is unique. having 6 of these on your form is just dumb
    
    On Error GoTo Whoops1:
    Dim C   As Control
    Dim bFound As Boolean
    Dim pic As StdPicture
    '1)
    Set pic = UserControl.Parent.Picture
    '2)
    For Each C In Parent.Controls
        If TypeOf C Is Duncan_GradientBackground Then
            If bFound = False Then
                bFound = True
            Else
                'one had already been found
                'so disable this one
                GoTo Whoops2
            End If
        End If
    Next C
    
    m_Validated = True
    ValidateControl = True
    
    Exit Function
Whoops1:
    ValidateControl = False
    Debug.Print UserControl.Parent.Name & " not suitable parent control"
    
    Exit Function
Whoops2:
    ValidateControl = False
    Debug.Print "Only one version of Duncan_GradientBackground is needed per form"
    
End Function


'------------
'USER CONTROL
'------------
Private Sub UserControl_InitProperties()
    ValidateControl

    'set default colours
    m_ColourTop = RGB(252, 252, 254)       'whiteish
    m_ColourBottom = RGB(203, 225, 252)    'light blue

End Sub

Private Sub UserControl_Paint()
    'The usercontrol has its InvisibleAtRuntime property set to true
    'so this is simply for when we are in design mode
    DrawTeaser
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ValidateControl
    If m_Validated Then
        If Ambient.UserMode Then
            Call Subclass_Start(UserControl.Parent.hwnd)
            Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_ENTERSIZEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_EXITSIZEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_NCPAINT, MSG_AFTER)
        End If
        
        With PropBag
            ColourTop = .ReadProperty("ColourTop", 0)
            ColourBottom = .ReadProperty("ColourBottom", 0)
        End With
    End If
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ColourTop", ColourTop
        .WriteProperty "ColourBottom", ColourBottom
    End With
End Sub

Private Sub UserControl_Resize()
    'fixed size
    Width = 50 * Screen.TwipsPerPixelX
    Height = 50 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Errs
    If Ambient.UserMode Then Call Subclass_StopAll
Errs:
End Sub


'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'END Subclassing Code===================================================================================


