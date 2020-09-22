VERSION 5.00
Begin VB.UserControl Splitter 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   ControlContainer=   -1  'True
   ScaleHeight     =   780
   ScaleWidth      =   1785
   ToolboxBitmap   =   "Splitter.ctx":0000
   Begin VB.PictureBox picMover 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   135
      ScaleHeight     =   105
      ScaleWidth      =   1410
      TabIndex        =   1
      Top             =   225
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   135
      ScaleHeight     =   105
      ScaleWidth      =   1410
      TabIndex        =   0
      Top             =   405
      Width           =   1410
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'  Product: Simple Splitter
'   Author: Lewis Miller (aka Deth)
'    Email: dethbomb@hotmail.com
'  Website: http://kickme.to/overdrive
'     Date: 12/10/03
'  Version: 1.1
'
' Changlog: Removed redundant code and variables, changed
'           movesplitterbar sub scope, to private...
'           fixed a bug in horizontal movement resizing
'           removed unnecasary code such as ScaleX etc...
'
'  Purpose: to simplify putting a splitter control on a form, and to
'           remove the need to distribute an ocx with your program
'
'Copyright: This code is copyright Lewis Miller December 10, 2003
'           You may not use this code in commercial applications
'           without express permission from the author (thats me).
'
'    Notes: When using horizontal mode, Set listbox's integeral height to false
'           to make resizing smoother. You should only use controls that can be
'           properly resized with in this control. An alternative is to use a
'           picture box if you wish to include several controls on one side
'           or the other. Please rport any bugs or changes to the psc comments
'           section, or on the forums at my site, thankyou...

'enum for orientation
Public Enum Splitter_Orientation
    Vertical = 0
    Horizontal = 1
End Enum

'Default Property Values:
Const m_def_SplitterWidth = 70
Const m_def_Orientation = 0
Const m_def_SplitterPosition = 50
Const m_def_SplitterMax = 90
Const m_def_SplitterMin = 10
Const m_def_SplitterColor = vbButtonFace

'Property Variables:
Dim m_SplitterWidth As Long
Dim m_LeftOrTopControl As Control
Dim m_Orientation As Splitter_Orientation
Dim m_RightOrBottomControl As Control
Dim m_SplitterPosition As Long
Dim m_SplitterMax As Long
Dim m_SplitterMin As Long
Dim m_SplitterColor As OLE_COLOR

'events for splitter
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event SplitComplete()
Event SplitMove()
Event SplitBegin()

'boolean yes/no value for calculating user mode
Dim blnRuntimeMode As Boolean

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'they clicked on splitter bar, so move mover bar to show there
    With picSplitter
        picMover.Move .Left, .Top, .Width, .Height
    End With
    
    'show it
    picMover.Visible = True
    picMover.ZOrder vbBringToFront

    RaiseEvent SplitBegin

End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim lngPosition As Long

    'dragging splitter bar, so we move the mover bar along with mouse trail to x or y
    With UserControl

        If m_Orientation = Horizontal Then
            
            lngPosition = Y + picSplitter.Top
            If lngPosition < Int((.Height * m_SplitterMax) / 100) And lngPosition > Int((.Height * m_SplitterMin) / 100) Then
                picMover.Move 0, Y + picSplitter.Top, .Width, m_SplitterWidth
                m_SplitterPosition = Int(100 / (.Height / lngPosition))
            End If

        Else

            lngPosition = X + picSplitter.Left
            If lngPosition < Int((.Width * m_SplitterMax) / 100) And lngPosition > Int((.Width * m_SplitterMin) / 100) Then
                picMover.Move X + picSplitter.Left, 0, m_SplitterWidth, .Height
                m_SplitterPosition = Int(100 / (.Width / lngPosition))
            End If

        End If

    End With

    RaiseEvent SplitMove

End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'done dragging, so hide mover bar, move splitter bar
    'to mover bar, and resize controls
    With picMover
        picSplitter.Move .Left, .Top, .Width, .Height
        .Visible = False
    End With

    Refresh

    RaiseEvent SplitComplete

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "Returns a handle (from Microsoft Windows) to the window a UserControl is contained in."

    ContainerHwnd = UserControl.ContainerHwnd

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    Enabled = picSplitter.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    picSplitter.Enabled() = New_Enabled

    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."

    hWnd = UserControl.hWnd

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  'this sub resizes any controls in the ocx, left/right or top/bottom
  
  On Error GoTo ErrHandle
  
    With picMover
        'move splitter bar to mover bar
        picSplitter.Move .Left, .Top, .Width, .Height
    End With
        
        'position controls
        'only do this if not in design mode
        If blnRuntimeMode Then

          With picSplitter
            If m_Orientation = Horizontal Then

                '*************************************************************************************
                'horizontal movement
                    If Not (m_LeftOrTopControl Is Nothing) Then

                        m_LeftOrTopControl.Move 0, 0, UserControl.ScaleWidth, .Top

                    End If

                    If Not (m_RightOrBottomControl Is Nothing) Then

                        m_RightOrBottomControl.Move 0, .Top + m_SplitterWidth, UserControl.ScaleWidth, (UserControl.ScaleHeight - .Top) - m_SplitterWidth

                    End If
                '******************************************************************************

              Else

                '******************************************************************************
                'vertical movement
                    If Not (m_LeftOrTopControl Is Nothing) Then

                        m_LeftOrTopControl.Move 0, 0, .Left, UserControl.Height

                    End If

                    If Not (m_RightOrBottomControl Is Nothing) Then

                        m_RightOrBottomControl.Move .Left + m_SplitterWidth, 0, (UserControl.ScaleWidth - .Left) + m_SplitterWidth, UserControl.ScaleHeight

                    End If

                '******************************************************************************
            End If
        End With
     End If


Exit Sub
ErrHandle:
            'most likely NOT a control that supports: width,height,top,left
            Err.Raise vbObjectError + Err.Number, UserControl.Ambient.DisplayName, Err.Description
End Sub

Private Sub UserControl_Resize()

    RaiseEvent Resize
    
    MoveSplitterBar
    Refresh

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get LeftOrTopControl() As Object
Attribute LeftOrTopControl.VB_Description = "Sets/ returns the top or left control of the splitter"

    Set LeftOrTopControl = m_LeftOrTopControl

End Property

Public Property Set LeftOrTopControl(New_LeftOrTopControl As Object)

    Set m_LeftOrTopControl = New_LeftOrTopControl
    PropertyChanged "LeftOrTopControl"
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Orientation() As Splitter_Orientation
Attribute Orientation.VB_Description = "Specifies what direction the splitter is set."

    Orientation = m_Orientation

End Property

Public Property Let Orientation(ByVal New_Orientation As Splitter_Orientation)
  
  'changes the up/down or top/bottom mode
  
    m_Orientation = New_Orientation
    
    '.Move is a built in vb method that simplifies setting the left, top, width, and height properties of a control
    'with this method you can do it all in one shot
    With picSplitter
        If m_Orientation = Horizontal Then 'up/down or top/bottom?
            .MousePointer = vbSizeNS 'change mouse
            .Move 0, Int((UserControl.Height * m_SplitterPosition) / 100), UserControl.Width, m_SplitterWidth
          Else
            .MousePointer = vbSizeWE 'change mouse
            .Move Int((UserControl.Width * m_SplitterPosition) / 100), 0, m_SplitterWidth, UserControl.Height
        End If

        picMover.Move .Left, .Top, .Width, .Height
    End With

    PropertyChanged "Orientation"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get RightOrBottomControl() As Object
Attribute RightOrBottomControl.VB_Description = "Returns/Sets the botoom, or top control."

    Set RightOrBottomControl = m_RightOrBottomControl

End Property

Public Property Set RightOrBottomControl(New_RightOrBottomControl As Object)

    Set m_RightOrBottomControl = New_RightOrBottomControl
    PropertyChanged "RightOrBottomControl"
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get SplitterWidth() As Long

    SplitterWidth = m_SplitterWidth

End Property

Public Property Let SplitterWidth(ByVal New_SplitterWidth As Long)

    m_SplitterWidth = New_SplitterWidth
    PropertyChanged "SplitterWidth"
    picSplitter.Width = m_SplitterWidth
    MoveSplitterBar
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,50
Public Property Get SplitterPosition() As Long
Attribute SplitterPosition.VB_Description = "Sets/Returns the splitter position."

    SplitterPosition = m_SplitterPosition

End Property

Public Property Let SplitterPosition(ByVal New_SplitterPosition As Long)

    If New_SplitterPosition > m_SplitterMax Then
        New_SplitterPosition = m_SplitterMax
    End If
    
    If New_SplitterPosition < m_SplitterMin Then
        New_SplitterPosition = m_SplitterMin
    End If
    
    m_SplitterPosition = New_SplitterPosition
    PropertyChanged "SplitterPosition"

    MoveSplitterBar
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,90
Public Property Get SplitterMax() As Long
Attribute SplitterMax.VB_Description = "Maximum Right or Top of Splitter Position"

    SplitterMax = m_SplitterMax

End Property

Public Property Let SplitterMax(ByVal New_SplitterMax As Long)

    If New_SplitterMax < 0 Then New_SplitterMax = 1
    If New_SplitterMax > 100 Then New_SplitterMax = 99
    
    If New_SplitterMax < m_SplitterMin Then
        m_SplitterMax = m_SplitterMin + 1
    End If
    
    m_SplitterMax = New_SplitterMax
    PropertyChanged "SplitterMax"

    MoveSplitterBar
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,10
Public Property Get SplitterMin() As Long
Attribute SplitterMin.VB_Description = "Minimum Left or Bottom of Splitter Position"

    SplitterMin = m_SplitterMin

End Property

Public Property Let SplitterMin(ByVal New_SplitterMin As Long)

    If m_SplitterMin > 100 Then m_SplitterMin = 99
    If m_SplitterMin < 0 Then m_SplitterMin = 1
    
    If m_SplitterMin > m_SplitterMax Then
        m_SplitterMin = m_SplitterMax - 1
    End If
    
    m_SplitterMin = New_SplitterMin
    PropertyChanged "SplitterMin"

    MoveSplitterBar
    Refresh

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,vbButtonFace
Public Property Get SplitterColor() As OLE_COLOR
Attribute SplitterColor.VB_Description = "Color Of The Splitter Bar"

    SplitterColor = m_SplitterColor

End Property

Public Property Let SplitterColor(ByVal New_SplitterColor As OLE_COLOR)

    m_SplitterColor = New_SplitterColor
    picSplitter.BackColor = New_SplitterColor
    PropertyChanged "SplitterColor"

End Property

Private Sub MoveSplitterBar()
    'this sub is a generic all purpose method to move the mover bar into correct position
    
    With UserControl
        If m_Orientation = Horizontal Then

            picMover.Move 0, Int((.Height * m_SplitterPosition) / 100), .Width, m_SplitterWidth

          Else
            
            picMover.Move Int((.Width * m_SplitterPosition) / 100), 0, m_SplitterWidth, .Height
        
        End If
    End With

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Orientation = m_def_Orientation
    m_SplitterPosition = m_def_SplitterPosition
    m_SplitterMax = m_def_SplitterMax
    m_SplitterMin = m_def_SplitterMin
    m_SplitterColor = m_def_SplitterColor
    m_SplitterWidth = m_def_SplitterWidth

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'determine runtime/design mode (we only move bar, and not controls in design mode)
    blnRuntimeMode = Ambient.UserMode

    picSplitter.Enabled = PropBag.ReadProperty("Enabled", True)

    Set m_LeftOrTopControl = PropBag.ReadProperty("LeftOrTopControl", Nothing)
    Set m_RightOrBottomControl = PropBag.ReadProperty("RightOrBottomControl", Nothing)
    
    SplitterColor = PropBag.ReadProperty("SplitterColor", m_def_SplitterColor)
    
    'the order of these is important
    SplitterWidth = PropBag.ReadProperty("SplitterWidth", m_def_SplitterWidth)
    SplitterMax = PropBag.ReadProperty("SplitterMax", m_def_SplitterMax)
    SplitterMin = PropBag.ReadProperty("SplitterMin", m_def_SplitterMin)
    SplitterPosition = PropBag.ReadProperty("SplitterPosition", m_def_SplitterPosition)

    Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)

    picSplitter.BackColor = m_SplitterColor
    picSplitter.Width = m_SplitterWidth

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", picSplitter.Enabled, True)
    Call PropBag.WriteProperty("LeftOrTopControl", m_LeftOrTopControl, Nothing)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("RightOrBottomControl", m_RightOrBottomControl, Nothing)
    Call PropBag.WriteProperty("SplitterPosition", m_SplitterPosition, m_def_SplitterPosition)
    Call PropBag.WriteProperty("SplitterMax", m_SplitterMax, m_def_SplitterMax)
    Call PropBag.WriteProperty("SplitterMin", m_SplitterMin, m_def_SplitterMin)
    Call PropBag.WriteProperty("SplitterColor", m_SplitterColor, m_def_SplitterColor)
    Call PropBag.WriteProperty("SplitterWidth", m_SplitterWidth, m_def_SplitterWidth)
    
End Sub

