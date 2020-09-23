VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API"
   ClientHeight    =   3165
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optNeither 
      Caption         =   "&Neither"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2970
      Width           =   960
   End
   Begin VB.OptionButton optPublic 
      Caption         =   "P&ublic"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   2790
      Width           =   960
   End
   Begin VB.OptionButton optPrivate 
      Caption         =   "&Private"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   2610
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.Timer Timer1 
      Left            =   90
      Top             =   0
   End
   Begin VB.ListBox lstTypes 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "frmAddIn.frx":0000
      Left            =   4635
      List            =   "frmAddIn.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1575
      Width           =   3525
   End
   Begin VB.ListBox lstConstants 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "frmAddIn.frx":0004
      Left            =   4590
      List            =   "frmAddIn.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   135
      Width           =   3525
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2565
      TabIndex        =   2
      Top             =   2565
      Width           =   1500
   End
   Begin VB.CommandButton cmdCopySelected 
      Caption         =   "Copy selected to clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1080
      TabIndex        =   1
      Top             =   2565
      Width           =   1500
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum enSelfTag
     tagApi = 0
     tagConstant = 1
     tagType = 2
End Enum

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public VBInstance                             As VBIDE.VBE
Public Connect                                  As Connect

Dim selfTag                                       As enSelfTag
Dim KeyToggleDown                         As Boolean
 
Private WithEvents cApiDeclares     As clsApiDeclares
Attribute cApiDeclares.VB_VarHelpID = -1
Private WithEvents cKey                  As clsKeys
Attribute cKey.VB_VarHelpID = -1



Private Sub Form_Load()
    ' load the constants list into lstConstants
   Call LoadConstantsList
   Call LoadTypesList
   ' always keeps this form on top of the vbinstance
   SetParent Me.hWnd, VBInstance.MainWindow.hWnd
   Set cApiDeclares = New clsApiDeclares
   Set cKey = New clsKeys
   cKey.Initialize Timer1
End Sub

Private Sub Form_Terminate()
   Set cKey = Nothing
   Set cApiDeclares = Nothing
End Sub

Private Sub cmdClose_Click()
  ' clear the main listbox
    lstMain.Clear
    Visible = False
End Sub







Private Sub cKey_LeftButtonAKeyDown()
       
       Dim currApi  As String, selectedWord As String

       selfTag = tagApi
       ' pass the word the user has selected on to see if it
       ' matches a valid api call in the list in clsApiDeclares
       selectedWord = ReturnHilightedWord
       cApiDeclares.ReturnedApiDeclare selectedWord
       
End Sub
Private Sub cApiDeclares_ApiToPaste(strApiToPaste As String)
 
  Dim strToCheck   As String
  ' we'll look for the api call up to the first  "
  strToCheck = "Private " + Left(strApiToPaste, InStr(1, strApiToPaste, Chr(34)))
  
  With VBInstance.ActiveCodePane.CodeModule
      ' make sure this api isnt already pasted
      If .Find(strToCheck, 1, 1, .CountOfDeclarationLines + 1, 500) = False Then
         '  paste the api returned by clsApiDeclares into the ide
           .InsertLines 2, "Private " & strApiToPaste
      End If
  End With
  
End Sub

Private Sub cKey_LeftButtonCKeyDown()

      Dim lcnt                   As Long
      Dim selectedWord   As String
      Dim left3                  As String
      Dim listMid3             As String


      selectedWord = ReturnHilightedWord

      selfTag = tagConstant

      ' if no word is selected then show all constants
      If Len(selectedWord) = 0 Then
           For lcnt = 0 To lstConstants.ListCount - 1
                lstMain.AddItem lstConstants.List(lcnt)
           Next lcnt

      Else
            'get the left 3 chrs of the selected word
            left3 = LCase(Trim(Left(selectedWord, 3)))

            'go through lstConstants and look for constants whose
            'first 3 letters match the first 3 letters of selected word
            For lcnt = 0 To lstConstants.ListCount - 1
                listMid3 = LCase(Trim(Mid(lstConstants.List(lcnt), 7, 3)))

                If listMid3 = left3 Then
                     lstMain.AddItem lstConstants.List(lcnt)
                End If
            Next lcnt
      End If

       ' show our form if there is something in the main list
      If lstMain.ListCount > 0 Then Visible = True

End Sub

Private Sub cKey_LeftButtonTKeyDown()

      Dim lcnt                    As Long
      Dim selectedWord    As String
      Dim left3                   As String
      Dim listMid3              As String

      selectedWord = ReturnHilightedWord
      selfTag = tagType

       ' if no word is selected then show all Types
      If Len(selectedWord) = 0 Then
            For lcnt = 0 To lstTypes.ListCount - 1
                 lstMain.AddItem lstTypes.List(lcnt)
            Next lcnt

      Else
            'get the left 3 chrs of the selected word
            left3 = LCase(Trim(Left(selectedWord, 3)))

            ' look for type matches in lstTyped
            For lcnt = 0 To lstTypes.ListCount - 1
                listMid3 = LCase(Trim(Mid(lstTypes.List(lcnt), 6, 3)))

                If listMid3 = left3 Then
                     lstMain.AddItem lstTypes.List(lcnt)
                End If
            Next lcnt

      End If

       ' show our form if there is something in the main list
      If lstMain.ListCount > 0 Then Visible = True

End Sub
 


'-------------------------------------------------------------------------------
' this sub copies any items selected in lstMain
' to the system clipboard preceeded by either
' "Private " or "Public "
'--------------------------------------------------------------------------------
Private Sub cmdCopySelected_Click()
    
    Dim stemp As String, scope As String
    Dim lcnt As Long
    
    ' are we preceeding selected with private or public or neither
    If optPrivate.Value = True Then
        scope = "Private "
    ElseIf optPublic.Value = True Then
        scope = "Public "
    ElseIf optNeither.Value = True Then
        scope = ""
    End If
    
    ' copy selected items from list to clipboard
    With lstMain
        For lcnt = 0 To .ListCount - 1
            If .Selected(lcnt) Then
            
                '   were copying constants to clipboard
                If selfTag = tagConstant Then
                       stemp = (stemp + scope + .List(lcnt) + vbCrLf)
                       
                '   were copying types to clipboard
                ElseIf selfTag = tagType Then
                       Dim bOn As Boolean
                       Dim ffile As Integer
                       Dim strHolder As String
                       Dim spaceHolder As String
                       Dim scopeVal      As String
                       
                       ffile = FreeFile
                        '  open the api list and look for the Types that match
                        '  the types selected in the listbox
                       Open funcTextPath For Input As #ffile
                           Do Until EOF(ffile)
                              Input #ffile, strHolder
                              ' were only interested in examining types
                              If Trim(Left(strHolder, 4)) = "Type" Or bOn Then
                                  ' an item selected in the list matches an item in the api textfile
                                  If Trim(strHolder) = .List(lcnt) Then
                                       bOn = True
                                  ElseIf Trim(Left(strHolder, 8)) = "End Type" Then
                                       stemp = (stemp + strHolder + vbCrLf)
                                       bOn = False
                                  End If
          
                                  If bOn Then
                                        If Trim(strHolder) = .List(lcnt) Then
                                                spaceHolder = ""
                                                scopeVal = scope
                                        Else
                                                spaceHolder = "     "
                                                scopeVal = ""
                                        End If
                                        
                                        stemp = (stemp + scopeVal + spaceHolder + strHolder + vbCrLf)
                                   End If
                              End If
                          Loop
                  End If
            End If
        Next lcnt
     End With
    
    ' clipboard clear + copying action
    Clipboard.Clear
    Clipboard.SetText stemp
    
End Sub
 

Private Function ReturnHilightedWord() As String

        ' this function returns the word that is selected in the vb ide
        Dim a&, b&, c&, d&
        Dim currline As String, currApi As String
         
        ' get the line of code currently selected
        VBInstance.ActiveCodePane.GetSelection a, b, c, d
        currline = VBInstance.ActiveCodePane.CodeModule.Lines(a, 1)
        ' returns the exact letters selected
        ReturnHilightedWord = LCase$(Trim$(Mid$(currline, b, (d - b))))
        
End Function
  

'--------------------------------------------------------------------------------------------------------------
' this sub loads all the items in microsofts original api list
' that start with the word "Const"..which means they are
' the constants that we are after
'---------------------------------------------------------------------------------------------------------------
Private Sub LoadConstantsList()
    Call LoadTextFile(True)
End Sub
 
'-----------------------------------------------------------------------------------------------
' this sub loads all the items in microsofts original api list
' that start with the word "Type"..which means they are
' the types that we are after (no pun intended)
'-----------------------------------------------------------------------------------------------
Private Sub LoadTypesList()
     Call LoadTextFile(False)
End Sub

Private Sub LoadTextFile(bConstants As Boolean)

     Dim ffile As Integer, chrLen As Integer, stemp As String, strVal As String
     Dim lst As ListBox
     
     ffile = FreeFile
      
     ' if were loading the list of constants from "\API.txt" into lstConstants...
     If bConstants = True Then
          strVal = "Const"
          chrLen = 5
          Set lst = lstConstants
     ' if were loading the list of types from "\API.txt" into lstTypes...
     Else
          strVal = "Type"
          chrLen = 4
          Set lst = lstTypes
     End If
     
     Open funcTextPath For Input As #ffile
         Do Until EOF(ffile)
              Input #ffile, stemp
           
              If Trim(Left(stemp, chrLen)) = strVal Then
                    lst.AddItem stemp
              End If
         Loop
    Close #ffile
    
    Set lst = Nothing
    
End Sub

Private Function funcTextPath() As String
     funcTextPath = App.Path + "\API.txt"
End Function
 
 
