VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const VK_A                As Long = 65
Private Const VK_C                As Long = 67
Private Const VK_T                As Long = 84
Private Const VK_LBUTTON  As Long = &H1
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim m_b_keyToggle                     As Boolean
Private WithEvents m_timer     As Timer

Event LeftButtonAKeyDown()
Event LeftButtonCKeyDown()
Event LeftButtonTKeyDown()


Public Sub Initialize(yourTimer As Timer)
   Set m_timer = yourTimer
   With m_timer
                .Interval = 50
                .Enabled = True
    End With
End Sub
Private Sub Class_Terminate()
   m_timer.Enabled = False
   Set m_timer = Nothing
End Sub

' this function returns true if the   a   key is being pressed
Private Function LeftButtonAKeyDown() As Boolean

   Dim ret As Long
   ret = GetKeyState(VK_A)
   If ret = -127 Or ret = -128 Then LeftButtonAKeyDown = True
  
End Function

' this function returns true if the   c   key is being pressed
Private Function LeftButtonCKeyDown() As Boolean

   Dim ret As Long
   ret = GetKeyState(VK_C)
   If ret = -127 Or ret = -128 Then LeftButtonCKeyDown = True
   
End Function

' this function returns true if the   t   key is being pressed
Private Function LeftButtonTKeyDown() As Boolean
   
   Dim ret As Long
   ret = GetKeyState(VK_T)
   If ret = -127 Or ret = -128 Then LeftButtonTKeyDown = True
   
End Function
   
Private Sub m_timer_Timer()
   
   Dim ret As Long
   ret = GetKeyState(VK_LBUTTON)
  
     ' if the left mouse is down...
   If ret = -127 Or ret = -128 Then
           ' if the  a  key is down...
          If LeftButtonAKeyDown = True Then
             If Not (m_b_keyToggle) Then
                  m_b_keyToggle = True
                  RaiseEvent LeftButtonAKeyDown
              End If
              
            ' if the  c  key is down ......
          ElseIf LeftButtonCKeyDown Then
              If Not (m_b_keyToggle) Then
                  m_b_keyToggle = True
                  RaiseEvent LeftButtonCKeyDown
              End If
            
            ' if the  t  key is down....
          ElseIf LeftButtonTKeyDown Then
              If Not (m_b_keyToggle) Then
                  m_b_keyToggle = True
                  RaiseEvent LeftButtonTKeyDown
              End If
              
          Else
             ' neither   a   or  c   or   t  is currently being pressed
              m_b_keyToggle = False
              
  End If
End If
 
   
End Sub
