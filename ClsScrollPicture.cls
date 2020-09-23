VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsScrollPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Modified version of this class from my LineArt and Edge detection upload
'some of this is copyright Min Thant Sin
'based on code found at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41164&lngWId=1
'Modifications Copyright 2002 Roger Gilchrist rojagilkrist@hotmail.com
'converted to a class demonstrating use of WithEvents
'
'DESCRIPTION
'This class allows you to set a PictureBox containing a picture larger than PictureBox size
'such that you can scroll around the image using Scroll bars, mouse or Arrow keys
'
'REQUIRES: a PictureBox contained within another PictureBox, a VScrollBar and a HScrollBar
'
'Place in Declaration section of a Form/Module
'
'   Private|Public ScrollPic As New ClsScrollPicture
'
'Place in Form_Load/ Sub Main
'
'       ScrollPic.AssignControls PicBox, VScroll, HScroll
'OPTIONALLY Place this just below AssignControls call,
'           if you have two PictureBoxes which are linked.
'           It allows you to interlock the scrollbars of the two systems.
'       ScrollPic.AssignInterlockScrolls Vscrool2, HScroll2
'
'
'PROPERTIES these have Default values you probably don't need to change but if you want you can
''Scroll_Percent DEF = 0.1 set LargeChange value for scroll bars to a percentage of image size
''Move_Speed     DEF = 100 size of steps used by arrow keys
'
'That's it!
'
'Private Routines (Description only)
'KeepInBounds    Keeps a value between Min and Max values
'
'Because the class uses WithEvents all the behaviour is automatically executed
''If you need to create any other activity remember that the On-Form Event
''fires before the Class Event. So On-Form cannot use Class generated response data
Private m_Scroll_Percent            As Single
Private m_Move_Speed                As Long
Private LastMouseX                  As Single 'self-explanatory
Private LastMouseY                  As Single '           "
Private tmpLeft                     As Long
Private tmpTop                      As Long
Private XDiff                       As Long    'width difference
Private YDiff                       As Long    'height difference
Private HasInterlock                As Boolean
Private InterlockVScrl              As VScrollBar
Attribute InterlockVScrl.VB_VarHelpID = -1
Private InterlockHScrl              As HScrollBar
Attribute InterlockHScrl.VB_VarHelpID = -1
Private WithEvents M_MovingPic      As PictureBox
Attribute M_MovingPic.VB_VarHelpID = -1
'This is the picturebox you load to/draw on whatever
Private M_ContainerPic              As PictureBox
'M_ContainerPic is used to set the on-screen size of the picture
'M_ContainerPic does not need to be assigned as the
'Sub AssignControls uses M_MovingPic.Container to get it
Private WithEvents VScrl            As VScrollBar
Attribute VScrl.VB_VarHelpID = -1
Private WithEvents HScrl            As HScrollBar
Attribute HScrl.VB_VarHelpID = -1
Private WithEvents m_PicForm        As Form
Attribute m_PicForm.VB_VarHelpID = -1

'm_PicForm is the form on which the pictureboxes are located
'It is needed to allow KeyPreview to be set and the Arrow keys to scroll
'm_PicForm does not need to be assigned as the
'Sub AssignControls uses M_MovingPic.Parent to get it
Public Sub AssignControls(MovePic As PictureBox, _
                          VS As VScrollBar, _
                          HS As HScrollBar)

  'Set the Various controls that the class manipulates
  'NOTE M_ContainerPic and m_PicForm must exist but are not assigned
  '.Container and .Parent find them automatically
  'You could add further error checking to this to make sure that .Container exists and is a PictureBox
  '.Parent must exist; its a Fform.

  Set M_MovingPic = MovePic
  With M_MovingPic
    .AutoSize = True
    Set M_ContainerPic = .Container
    .Move 0, 0, M_ContainerPic.Width, M_ContainerPic.Height
    Set m_PicForm = .Parent
    m_PicForm.KeyPreview = True
  End With
  Set VScrl = VS
  Set HScrl = HS
  SetMoveValues

End Sub

Private Sub Class_Initialize()

  'Set Default Values

  m_Move_Speed = 100
  m_Scroll_Percent = 0.1

End Sub

Private Sub HScrl_Change()

  'WithEvents causes this to be called

  M_MovingPic.Left = (-HScrl.Value)
  tmpLeft = M_MovingPic.Left
  If HasInterlock Then
    InterlockHScrl.Value = HScrl.Value
  End If

End Sub

Private Sub HScrl_Scroll()

  'WithEvents causes this to be called

  HScrl_Change

End Sub

Private Sub KeepInBounds(ByVal Min As Long, _
                         lngVal As Long, _
                         ByVal Max As Long)

  'keeps Val between Min and Max

  If lngVal >= Max Then
    lngVal = Max
  End If
  If lngVal <= Min Then
    lngVal = Min
  End If

End Sub

Private Sub M_MovingPic_Change()

  SetMoveValues

End Sub

Private Sub M_MovingPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single) '

  'WithEvents causes this to be called
  'sets initial postion of mouse grag for mouse scrolling

  If Button = vbLeftButton Then
    LastMouseX = X
    LastMouseY = Y
  End If

End Sub

Private Sub M_MovingPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

  'WithEvents causes this to be called
  'Drags the image

  If Button = vbLeftButton Then
    If HScrl.Enabled Then
      tmpLeft = M_MovingPic.Left + (X - LastMouseX)
    End If
    If VScrl.Enabled Then
      tmpTop = M_MovingPic.Top + (Y - LastMouseY)
    End If
    'Make sure M_MovingPic doesn't go out of boundaries
    KeepInBounds -XDiff, tmpLeft, 0
    KeepInBounds -YDiff, tmpTop, 0
    'M_MovingPic.Move tmpLeft, tmpTop
    If VScrl.Enabled Then
      M_MovingPic.Top = tmpTop
    End If
    If HScrl.Enabled Then
      M_MovingPic.Left = tmpLeft
    End If
    'Change the scrollbar values
    'note the Abs
    If HScrl.Enabled Then
      HScrl.Value = Abs(tmpLeft)
    End If
    If VScrl.Enabled Then
      VScrl.Value = Abs(tmpTop)
    End If
  End If

End Sub

Private Sub m_PicForm_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

  'WithEvents causes this to be called
  'Note  negative inversion of XDiff, YDiff values

  Select Case KeyCode
   Case vbKeyLeft
    If HScrl.Enabled Then
      tmpLeft = tmpLeft + m_Move_Speed
    End If
   Case vbKeyRight
    If HScrl.Enabled Then
      tmpLeft = tmpLeft - m_Move_Speed
    End If
   Case vbKeyUp
    If VScrl.Enabled Then
      tmpTop = tmpTop + m_Move_Speed
    End If
   Case vbKeyDown
    If VScrl.Enabled Then
      tmpTop = tmpTop - m_Move_Speed
    End If
  End Select
  KeepInBounds -XDiff, tmpLeft, 0
  KeepInBounds -YDiff, tmpTop, 0
  If VScrl.Enabled Then
    M_MovingPic.Top = tmpTop
  End If
  If HScrl.Enabled Then
    M_MovingPic.Left = tmpLeft
  End If
  '    M_MovingPic.Move tmpLeft, tmpTop
  If HScrl.Enabled Then
    HScrl.Value = Abs(tmpLeft)
  End If
  If VScrl.Enabled Then
    VScrl.Value = Abs(tmpTop)
  End If

End Sub

Public Property Get Move_Speed() As Long

  'Allows yout Get size of steps used by arrow keys DEF = 100

  Move_Speed = m_Move_Speed

End Property

Public Property Let Move_Speed(ByVal Value As Long)

  'Allows yout Let size of steps used by arrow keys DEF = 100

  m_Move_Speed = Value

End Property

Public Property Get Scroll_Percent() As Single

  'Allows you to Get LargeChange value for scroll bars to a percentage of image size DEF = 0.1

  Scroll_Percent = m_Scroll_Percent

End Property

Public Property Let Scroll_Percent(ByVal Value As Single)

  'Allows you to Let LargeChange value for scroll bars to a percentage of image size DEF = 0.1

  m_Scroll_Percent = Value

End Property

Public Sub SetMoveValues()

  'call after loading a picture
  'THis routine resets picturebox sizes and Scroll bar values

  With M_MovingPic
    .Move 0, 0 ' put picture at Top, Left = 0, 0
    YDiff = Abs(.Height - M_ContainerPic.Height)
    XDiff = Abs(.Width - M_ContainerPic.Width)
    VScrl.Value = 0
    HScrl.Value = 0
    If .Width < M_ContainerPic.Width Then
      .Left = (M_ContainerPic.Width - .Width) / 2
    End If
    If .Height < M_ContainerPic.Height Then
      .Top = (M_ContainerPic.Height - .Height) / 2
    End If
    .Refresh
  End With
  With VScrl
    If M_MovingPic.Height > M_ContainerPic.Height Then
      .Max = YDiff
      .LargeChange = IIf(YDiff * m_Scroll_Percent > 0, YDiff * m_Scroll_Percent, 1)
      .Enabled = True
     Else
      .Max = 1
      .LargeChange = 1
      .Enabled = False
    End If
  End With
  With HScrl
    If M_MovingPic.Width > M_ContainerPic.Width Then
      .Max = XDiff
      .LargeChange = IIf(XDiff * m_Scroll_Percent > 0, XDiff * m_Scroll_Percent, 1)
      .Enabled = True
     Else
      .Max = 1
      .LargeChange = 1
      .Enabled = False
    End If
  End With

End Sub

Private Sub VScrl_Change()

  'WithEvents causes this to be called

  M_MovingPic.Top = (-VScrl.Value)
  tmpTop = M_MovingPic.Top
  If HasInterlock Then
    InterlockVScrl.Value = VScrl.Value
  End If

End Sub

Private Sub VScrl_Scroll()

  'WithEvents causes this to be called

  VScrl_Change

End Sub

'Public Sub AssignInterlockScrolls(VS As VScrollBar,
''                                  HS As HScrollBar)
'
'  'Optional Routine
'  'Allows you to create pairs of scrollable PictureBoxes which are interlocked
'  '(a move in one is reflected in the other) by linking their respective scrollbars
'
'  HasInterlock = True '    let rest of program know if this has been activated
'  Set InterlockVScrl = VS
'  Set InterlockHScrl = HS
'
'End Sub

':)Code Fixer V2.9.2 (1/02/2005 4:13:54 PM) 59 + 273 = 332 Lines Thanks Ulli for inspiration and lots of code.

