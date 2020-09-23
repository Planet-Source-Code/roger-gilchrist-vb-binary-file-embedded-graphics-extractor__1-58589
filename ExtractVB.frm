VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtractVB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract images from VB binary support files"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12270
   Icon            =   "ExtractVB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   12270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDisplay 
      Height          =   6000
      Left            =   6240
      TabIndex        =   7
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox picContainer 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   5505
         Left            =   120
         ScaleHeight     =   5505
         ScaleWidth      =   5595
         TabIndex        =   10
         Top             =   240
         Width           =   5595
         Begin VB.PictureBox picDisplay 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   3210
            Left            =   720
            ScaleHeight     =   3210
            ScaleWidth      =   3975
            TabIndex        =   11
            Top             =   840
            Width           =   3975
         End
      End
      Begin VB.VScrollBar vscDisplay 
         Enabled         =   0   'False
         Height          =   5505
         LargeChange     =   500
         Left            =   5760
         SmallChange     =   150
         TabIndex        =   9
         Top             =   225
         Width           =   150
      End
      Begin VB.HScrollBar hscDisplay 
         Enabled         =   0   'False
         Height          =   150
         LargeChange     =   500
         Left            =   -30
         TabIndex        =   8
         Top             =   5760
         Width           =   5835
      End
   End
   Begin MSComDlg.CommonDialog cdlExtract 
      Left            =   3960
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraExtractVB 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6195
      Begin VB.PictureBox picCFXPBugFixfrmExtractVB 
         BorderStyle     =   0  'None
         Height          =   4680
         Left            =   120
         ScaleHeight     =   4680
         ScaleWidth      =   6000
         TabIndex        =   1
         Top             =   175
         Width           =   6000
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete >"
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            TabIndex        =   19
            ToolTipText     =   "Delete the file displayed in the main image."
            Top             =   2760
            Width           =   885
         End
         Begin VB.CommandButton cmdAbort 
            Cancel          =   -1  'True
            Caption         =   "Abort"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   18
            Top             =   2760
            Width           =   885
         End
         Begin VB.PictureBox picGuage 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            DrawMode        =   7  'Invert
            FillColor       =   &H000000FF&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   0
            ScaleHeight     =   225
            ScaleWidth      =   5880
            TabIndex        =   13
            Top             =   4395
            Width           =   5940
         End
         Begin VB.CommandButton cmdStartExtraction 
            Caption         =   "Start Extraction"
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   12
            Top             =   2760
            Width           =   2085
         End
         Begin VB.CommandButton cmdExtractFrom 
            Caption         =   "Extract From..."
            Height          =   285
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Select frm, dob and ctl files to search for images."
            Top             =   0
            Width           =   1725
         End
         Begin VB.CommandButton cmdChangeDestination 
            Caption         =   "Change Destination folder..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   4
            ToolTipText     =   "By default files are created in source file folder"
            Top             =   1800
            Width           =   2565
         End
         Begin VB.TextBox txtDestination 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   2070
            Width           =   5985
         End
         Begin VB.ListBox lstSource 
            Height          =   1230
            Left            =   5
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   250
            Width           =   5985
         End
         Begin VB.PictureBox picHolder 
            Height          =   720
            Left            =   0
            ScaleHeight     =   660
            ScaleWidth      =   5880
            TabIndex        =   14
            Top             =   3360
            Width           =   5940
            Begin VB.HScrollBar HScroll 
               Enabled         =   0   'False
               Height          =   240
               Left            =   0
               TabIndex        =   16
               Top             =   420
               Width           =   5880
            End
            Begin VB.PictureBox picThumbnail 
               BorderStyle     =   0  'None
               Height          =   420
               Left            =   0
               ScaleHeight     =   420
               ScaleWidth      =   420
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   0
               Width           =   420
               Begin VB.Shape shpSelector 
                  BorderColor     =   &H000080FF&
                  BorderStyle     =   2  'Dash
                  BorderWidth     =   5
                  Height          =   495
                  Left            =   0
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.Image imgThumbnail 
                  Height          =   420
                  Index           =   0
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   420
               End
            End
         End
         Begin VB.Label lblExtractVB 
            Alignment       =   2  'Center
            Height          =   210
            Left            =   -600
            TabIndex        =   17
            Top             =   4140
            UseMnemonic     =   0   'False
            Width           =   2820
         End
      End
   End
   Begin VB.Label lblFileDescription 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   5040
      Width           =   6135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmExtractVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bProcessing                 As Boolean
Private bAbort                      As Boolean
Private sSource()                   As String
Private nImageCount                 As Long
Private CurImage                    As Long
Private Type ImageData
  ipath                             As String
  iName                             As String
  iExt                              As String
  iKB                               As Long
  iHieght                           As Long
  iWidth                            As Long
  iType                             As Long
  iThumbID                          As Long
End Type
Private PicData()                   As ImageData
Private ScrollPic                   As New ClsScrollPicture
' ListBox Tooltips control
Private Const LB_ITEMFROMPOINT      As Long = &H1A9
Private strBaseCaption              As String
Private Declare Function SendLBMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                          ByVal wMsg As Long, _
                                                                          ByVal wParam As Long, _
                                                                          lParam As Any) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub CaptionComment(ByVal strCom As String)

  Caption = strBaseCaption & IIf(Len(strCom), "...", vbNullString) & strCom

End Sub

Private Sub ClearForExtraction()

  Dim I As Long

  bAbort = False
  cmdStartExtraction.Enabled = False
  cmdAbort.Enabled = True
  ProgressBar 0
  For I = imgThumbnail.Count - 1 To 1 Step -1
    Unload imgThumbnail(I)
  Next I
  imgThumbnail(0).Visible = True
  picThumbnail.Width = 1
  HScroll.Enabled = False
  lblExtractVB.Caption = "No images extracted"
  Erase PicData

End Sub

Private Sub cmdAbort_Click()

  If bProcessing Then
    bAbort = True
    DoEvents
   Else
    Unload Me
  End If

End Sub

Private Sub cmdChangeDestination_Click()

  Dim sFolder As String

  sFolder = FolderBrowser("Select destination folder for the images:", Me.hWnd)
  If LenB(sFolder) Then
    txtDestination.Text = sFolder
    Set_OK_State
  End If

End Sub

Private Sub cmdDelete_Click()

  On Error Resume Next
  If LenB(PicData(CurImage).iName) Then
    cmdDelete.Enabled = False
    Kill PicData(CurImage).ipath & "\" & PicData(CurImage).iName
    With PicData(CurImage)
      .iExt = ""
      .iHieght = 0
      .iKB = 0
      .iName = ""
      .ipath = ""
      .iType = 0
      .iWidth = 0
      If .iThumbID > 0 Then
        Unload imgThumbnail(.iThumbID)
       Else
        imgThumbnail(.iThumbID).Visible = False
      End If
    End With
    PositionThumbs
    lblFileDescription.Caption = ""
    picDisplay = LoadPicture()
    If CurImage > 0 And CurImage < imgThumbnail.Count Then
      imgThumbnail_Click CInt(CurImage + 1)
     Else
      If CurImage > 1 Then
        imgThumbnail_Click CInt(CurImage - 1)
      End If
    End If
  End If
  On Error GoTo 0

End Sub

Private Sub cmdExtractFrom_Click()

  Dim n         As Long
  Dim sFolder   As String
  Dim I         As Long
  Dim sFileName As String
  Dim nCount    As Long

  On Error GoTo PickSourceCancelled
  With cdlExtract
    .DialogTitle = "Open VB files"
    'Fixed thanks Tony
    .Filter = "VB binary support files(*.frx;*.dox;*.ctx;*.dsx;*.pax)|*.frx;*.dox;*.ctx;*.dsx;*.pax"
    .FilterIndex = 1
    .CancelError = True
    .Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer
    .FileName = ""
    .MaxFileSize = 5120
    .ShowOpen
    .CancelError = False
    sFileName = .FileName
  End With
  If LenB(sFileName) Then
    ' Build sSource() array
    nCount = 0
    Erase sSource
    cmdChangeDestination.Enabled = True
    n = InStr(sFileName, vbNullChar)
    If n > 0 Then   ' Multi-select
      ' First one is the folder
      sFolder = Left$(sFileName, n - 1)
      txtDestination.Text = sFolder
      sFileName = Mid$(sFileName, n + 1)
      ' The rest are the files
      Do While n > 0
        n = InStr(sFileName, vbNullChar)
        ReDim Preserve sSource(0 To nCount)
        If n = 0 Then
          sSource(nCount) = AttachPath(sFileName, sFolder)
         Else
          sSource(nCount) = AttachPath(Left$(sFileName, n - 1), sFolder)
          sFileName = Mid$(sFileName, n + 1)
        End If
        nCount = nCount + 1
      Loop
     Else            ' Single file...
      ReDim sSource(0)
      sSource(0) = sFileName
      txtDestination.Text = Left$(sFileName, InStrRev(sFileName, "\"))
      nCount = 1
    End If
    ' Fill listbox
    With lstSource
      .Clear
      For I = 0 To (nCount - 1)
        If SourceFileExists(sSource(I)) Then
          'Fixed this stops binary files loading if the main form is missing Thanks Tony
          .AddItem ExtractFileName(sSource(I))
          .ItemData(.NewIndex) = I
         Else
          MsgBox "The file '" & ConvertXfileToMainFile(ExtractFileName(sSource(I))) & "' is missing so the binary file will not be loaded."
          sSource(I) = ""
        End If
      Next I
    End With
    Set_OK_State
  End If
PickSourceCancelled:
  cdlExtract.CancelError = False

End Sub

Private Function SourceFileExists(strFname As String) As Boolean


  SourceFileExists = FileExist(ConvertXfileToMainFile(strFname))

End Function

Private Sub cmdStartExtraction_Click()

  ExtractImages
  cmdDelete.Enabled = True

End Sub

Private Function ConvertXfileToMainFile(varFile As Variant) As String

  '*.frm;*.dob;*.ctl;*.dsr;*pag

  Select Case LCase$(Right$(varFile, 4))
   Case ".frx"
    ConvertXfileToMainFile = Left$(varFile, Len(varFile) - 3) & "frm"
   Case ".dox"
    ConvertXfileToMainFile = Left$(varFile, Len(varFile) - 3) & "dob"
   Case ".ctx"
    ConvertXfileToMainFile = Left$(varFile, Len(varFile) - 3) & "ctl"
   Case ".drx"
    ConvertXfileToMainFile = Left$(varFile, Len(varFile) - 3) & "dsr"
   Case ".pax"
    ConvertXfileToMainFile = Left$(varFile, Len(varFile) - 3) & "pag"
  End Select

End Function

Private Sub DisplayPicData(ByVal picID As Long)

  With PicData(picID)
    lblFileDescription.Caption = "Path: " & .ipath & vbNewLine & _
                                 "Name: " & .iName & vbNewLine & _
                                 "Size: " & .iKB & "KB   Height: " & .iHieght & "     Width: " & .iWidth & "   Type: " & .iType
  End With

End Sub

' Icon = "FormFile.frx":0000
'      ^               ^     = Markers
'        |-----------------| = Parameter
'
' Returns the image data in a string
'
Private Function ExtractImage(ByVal sString As String, _
                              sSourceFile As String, _
                              PrevOffset As Long) As String

  Dim nHandle   As Long
  Dim nOffset   As Long
  Dim nFileSize As Long
  Dim nSize     As Long
  Dim sFile     As String
  Dim sData     As String
  Dim sBytes    As String
  Dim bFileOpen As Boolean
  Dim n         As Long

  On Error GoTo EI_ErrorHandler
  n = InStr(sString, ":")
  If n Then
    sFile = AttachPath(StripQuotes(Left$(sString, n - 1)), ExtractPath(sSourceFile))
    If FileExist(sFile) Then
      sString = "&H" & Trim$(Mid$(sString, n + 1))
      nOffset = CLng(sString) + 1 '+ PrevOffset
      PrevOffset = nOffset - 1
      nHandle = FreeFile
      Open sFile For Binary Access Read Shared As #nHandle
      bFileOpen = True
      nFileSize = LOF(nHandle)
      If (nOffset + 12) > nFileSize Then
        GoTo EI_ErrorHandler
      End If
      ' Get the header...
      Seek #nHandle, nOffset
      sData = Mid$(Input$(12, #nHandle), 9, 4)
      ' Byte 9 to 12 (long) contains data size
      sBytes = "&H" & Right$("00" & Hex$(Asc(Mid$(sData, 4, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 3, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 2, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 1, 1))), 2)
      nSize = CLng(sBytes)
      If nSize < 0 Or (nOffset + 11 + nSize) > nFileSize Then
        ' Try 28 byte header
        If (nOffset + 27) > nFileSize Then
          GoTo EI_ErrorHandler
        End If
        ' Get the header...
        Seek #nHandle, nOffset
        sData = Mid$(Input$(28, #nHandle), 25, 4)
        ' Byte 25 to 28 (long) contains data size
        sBytes = "&H" & Right$("00" & Hex$(Asc(Mid$(sData, 4, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 3, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 2, 1))), 2) & Right$("00" & Hex$(Asc(Mid$(sData, 1, 1))), 2)
        nSize = CLng(sBytes)
        If nSize < 0 Or (nOffset + 27 + nSize) > nFileSize Then
          GoTo EI_ErrorHandler
        End If
      End If
      ' Get the data (position: nOffset + 13 - Already in position)
      ExtractImage = Input$(nSize, #nHandle)
      ' That's it, the icon data is obtained
      Close #nHandle
      bFileOpen = False
      '   Else
      ''Fixed not needed any more thanks Tony
      'MsgBox "The file requires an FRX, DOX or CTX file but it is missing"
    End If
    Exit Function
EI_ErrorHandler:
    If bFileOpen Then
      Close #nHandle
    End If
  End If

End Function

Private Sub ExtractImages()

  
  Dim I                  As Long
  Dim J                  As Long
  Dim K                  As Long
  Dim sFileIn()          As String
  Dim nTotalSize         As Long    ' Total bytes to analyse (all files)
  Dim nReadSize          As Long
  Dim nProgress          As Long
  Dim nCount             As Long
  Dim nInCount           As Long
  Dim strFormName        As String
  Dim strControlName     As String
  Dim strControlIndex    As String
  Dim strControlProperty As String
  Dim sFolder            As String
  Dim sString            As String
  Dim sImageData         As String
  Dim arrHidden          As Variant
  Dim strFindIndex       As String

  'Dim sImageExt  As String
  'Dim n          As Long
  ''Dim bScan      As Boolean
  ClearForExtraction
  On Error GoTo ExtractError
  CaptionComment "Checking source..."
  sFolder = txtDestination.Text
  nCount = UBound(sSource)
  nInCount = 0
  nReadSize = 0
  nImageCount = 0
  'bScan = False
  ' Check of all files are available
  For I = 0 To nCount
    If FileExist(sSource(I)) Then
      ReDim Preserve sFileIn(0 To nInCount)
      sFileIn(nInCount) = ConvertXfileToMainFile(sSource(I))
      nInCount = nInCount + 1
      nTotalSize = nTotalSize + FileLen(sSource(I))
    End If
  Next I
  If bAbort Then
    GoTo ExtractExit
  End If
  If nInCount Then
    CaptionComment "Checking Target..."
    If FolderExist(sFolder) Then
      CaptionComment "Checks OK - Analysing"
      ' Yield to other processes - just in case Cancel is pressed
      DoEvents
      If bAbort Then
        GoTo ExtractExit
      End If
      For I = 0 To (nInCount - 1)
        ' Yield to other processes - just in case Cancel is pressed
        DoEvents
        If bAbort Then
          GoTo ExtractExit
        End If
        CaptionComment "Analysing " & ExtractFileName(sFileIn(I))
        sImageData = ""
        ' Open for for line-input...
        strFormName = ExtractFileName(sFileIn(I), False)
        GetHiddenTxt sFileIn(I), arrHidden
        For J = LBound(arrHidden) To UBound(arrHidden)
          ' Yield to other processes - just in case Cancel is pressed
          ' Update progressbar...
          nProgress = ((nReadSize + UBound(arrHidden)) * 100) / nTotalSize
          ProgressBar IIf(nProgress > 100, 100, nProgress)
          sString = arrHidden(J)
          If MatchString(sString, "BEGIN ") Then
            strControlName = Trim$(Mid$(sString, InStrRev(sString, " ")))
            'search for index for naming purposes
            'this has to be done because the properties are alpha-listed
            'so Down/DisabledPicture would be found before Index was set in Commandbuttons
            strControlIndex = ""
            K = J + 1
            Do
              strFindIndex = UCase$(arrHidden(K))
              If MatchString(strFindIndex, "INDEX ") Then
                strControlIndex = Trim$(Mid$(strFindIndex, InStrRev(strFindIndex, " ")))
                Exit Do
              End If
              K = K + 1
              'reached next object or end of data
            Loop Until MatchString(strFindIndex, "BEGIN ") Or K > UBound(arrHidden)
            If InStr(sString, "MSComctlLib.ImageList") Then
              ImgListExtract arrHidden, sFileIn(I), J, K - 1, sImageData, sFolder, strFormName, strControlName, strControlProperty, nImageCount
              J = K
            End If
           ElseIf IsFrxGraphicLine(sString, strControlProperty) Then
            sImageData = GetImageData(sString, sFileIn(I))
          End If
          'found an image so process it
          If LenB(sImageData) Then
            ProcessOneImage sImageData, sFolder, strFormName, strControlName, strControlIndex, strControlProperty, nImageCount
          End If
          'EndOfFileLoop:
          If bAbort Then
            Exit For
          End If
          DoEvents
        Next J
        nReadSize = nReadSize + UBound(arrHidden)
        If bAbort Then
          Exit For
        End If
      Next I
      ProgressBar 100
      CaptionComment "Extraction completed"
ExtractExit:
      On Error Resume Next
      cmdAbort.Enabled = False
      Set_OK_State
     Else
      CaptionComment "Invalid target folder"
      MsgBox "The target folder you specified is invalid. Please select another target folder.", vbExclamation, "Invalid Folder"
    End If
   Else
    CaptionComment "No files to analyse"
    MsgBox "There are no files to analyse. Please create a new list then try again.", vbExclamation, "No Files"
  End If
  CaptionComment ""

Exit Sub

ExtractError:
  MsgBox "Error occurred during extraction. Process aborted." & vbNewLine & _
       "(" & Err.Number & " - " & Err.Description & ")", vbCritical, "Extract Error"
  ProgressBar 0
  GoTo ExtractExit
  On Error GoTo 0

End Sub

Private Sub Form_Initialize()

  InitCommonControls

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

  If cmdDelete.Enabled Then
    If KeyCode = vbKeyDelete Then
      cmdDelete_Click
    End If
  End If
  If cmdAbort.Enabled Then
    If KeyCode = vbKeyEscape Then
      cmdAbort_Click
    End If
  End If

End Sub

Private Sub Form_Load()

  strBaseCaption = Caption
  ScrollPic.AssignControls picDisplay, vscDisplay, hscDisplay
  bProcessing = False
  bAbort = False
  nImageCount = 0

End Sub

Private Sub GenerateNewThumb(ByVal nImageCount As Long, _
                             ByVal sImageFile As String)

  On Error Resume Next
  If nImageCount = 0 Then
    imgThumbnail(0).Visible = True
   Else
    Load imgThumbnail(nImageCount)
  End If
  'picThumbnail.Width = 460 * nImageCount
  With imgThumbnail(nImageCount)
    '.Left = 460 * (nImageCount - 1)
    .Picture = picDisplay.Picture
    .ToolTipText = sImageFile
    .Visible = True
  End With
  PositionThumbs
  SelectFrame nImageCount
  On Error GoTo 0

End Sub

Private Sub GeneratePicData(ByVal strFile As String, _
                            strData As String, _
                            ByVal lngID As Long)

  'display the image so that the data can be gathered for PicData

  picDisplay = LoadPicture(strFile)
  ReDim Preserve PicData(lngID) As ImageData
  With PicData(lngID)
    .ipath = Left$(strFile, InStrRev(strFile, "\") - 1)
    .iName = Mid$(strFile, InStrRev(strFile, "\") + 1)
    .iExt = GetImageExtention(strData)
    .iKB = CLng(Len(strData) / 1024)
    .iHieght = ScaleY(picDisplay.Picture.Height)
    .iWidth = ScaleX(picDisplay.Picture.Width)
    .iType = picDisplay.Picture.Type
    .iThumbID = lngID
  End With

End Sub

Private Sub GetHiddenTxt(ByVal strFilename As String, _
                         ArrD As Variant)

  Dim FN      As Long
  Dim strData As String
  Dim strTemp As String

  FN = FreeFile
  Open strFilename For Input Access Read Shared As FN
  Do
    Line Input #FN, strTemp
    strData = strData & vbNewLine & Trim$(strTemp)
  Loop Until InStr(strTemp, "Attribute VB_")
  Close FN
  strData = Mid$(strData, 2)
  ArrD = Split(strData, vbNewLine)

End Sub

Private Function GetImageData(ByVal sString As String, _
                              strFilename As String, _
                              Optional PrevOffset As Long = 0) As String

  Dim n As Long

  n = InStr(sString, "=")
  If n Then
    sString = Trim$(Mid$(sString, n + 1))
    GetImageData = ExtractImage(sString, strFilename, PrevOffset)
  End If

End Function

Private Function GetImageExtention(ByVal sImageData As String) As String

  'bmp, gif, ico,jpg, wmf, cur

  If Left$(sImageData, 3) = "GIF" Then
    GetImageExtention = "gif"
   ElseIf Left$(sImageData, 2) = "BM" Then
    GetImageExtention = "bmp"
   ElseIf Left$(sImageData, 2) = (vbNullChar & vbNullChar) Then
    GetImageExtention = "ico" '.cur files are also recognised as ico
   ElseIf Mid$(sImageData, 7, 4) = "JFIF" Then
    GetImageExtention = "jpg" ' or jpeg or Tiff
   ElseIf Mid$(sImageData, 6, 5) = "¼Exif" Then
    GetImageExtention = "jpg" 'Or jpeg
   ElseIf Left$(sImageData, 4) = "Î-ãÜ" Then
    GetImageExtention = "wmf"
   ElseIf Mid$(sImageData, 42, 3) = "EMF" And Left$(sImageData, 1) = Chr$(1) Then
    'this is a bit of a fake I only had one emf file to experiment with ;)
    GetImageExtention = "emf"
  End If

End Function

Private Function GetStringValue(varCode As Variant) As String

  Dim arrTmp As Variant
  Dim strT   As String

  arrTmp = Split(varCode)
  strT = arrTmp(UBound(arrTmp))
  If strT = Chr$(34) & Chr$(34) Then
    GetStringValue = ""
   Else
    GetStringValue = Mid$(Left$(strT, Len(strT) - 1), 2)
  End If

End Function

Private Sub Guage(pic As Control, _
                  ByVal iPercent As Long)

  ' this routine will draw a 3D guage in the PictureBox control
  ' pic is the control
  ' iPercent% is the percentage to show in the guage
  ' this is useful if you want to only show the guage when something is
  ' happening but not show it at other times
  ' the percentage to show will be stored into the Tag property so that
  ' we can tell what it is currently set to if we need to repaint it at
  ' a random time
  
  Const XORPEN      As Long = 7
  Dim sPercent      As String
  Dim iLeft         As Long
  Dim iTop          As Long
  Dim iRight        As Long
  Dim iBottom       As Long
  Dim iLineWidth    As Long
  Const DGREYCOLOUR As Long = &H808080
  Const LGREYCOLOUR As Long = &HC0C0C0
  Const WHITECOLOUR As Long = &HFFFFFF
  Const COPYPEN     As Long = 13

  ' these are used to create the 3D effect
  ' validate our percentage
  If iPercent < 0 Then
    iPercent = 0
   ElseIf iPercent > 100 Then
    iPercent = 100
  End If
  ' set the number of twips per pixel into a variable
  ' NOTE: the picture control and the form it is on are expected to have
  ' their scale mode set to Twips
  iLineWidth = Screen.TwipsPerPixelX
  ' I leave the BorderStyle set to 1 at design time so that the control is
  ' easy to find, but at run time we want the border to be invisible,
  ' however, just switching the border off will actually trigger a refresh
  ' of the control which is no use if AutoRedraw is set to False because
  ' that will trigger this code to run which will trigger another refresh
  ' which will ...
  If pic.BorderStyle <> 0 Then
    pic.BorderStyle = 0
  End If
  ' save the percentage into the Tag property - we can use this to repaint
  ' the guage if AutoRedraw is set to False
  pic.Tag = iPercent
  ' set the text we will draw into a variable
  sPercent = CStr(iPercent) & "%"
  ' work out the co-ords for the percentage bar
  iLeft = iLineWidth
  iTop = iLineWidth
  iRight = pic.ScaleWidth - iLineWidth
  iBottom = pic.ScaleHeight - iLineWidth
  ' erase everything by redrawing the background
  With pic
    .DrawMode = COPYPEN
    pic.Line (iLeft, iTop)-(iRight, iBottom), pic.BackColor, BF
    ' add the text - work out where to put it first - nicely centered
    ' the default in VB3 is for bold text, change the FontBold property in
    ' the Picture control if you want this to be non-bold
    .CurrentX = (.ScaleWidth - .TextWidth(sPercent)) / 2
    .CurrentY = (.ScaleHeight - .TextHeight(sPercent)) / 2
    pic.Print sPercent
    ' do the two colour bar by setting the DrawMode XOr then draw the bar
    ' in the fillcolour, if this overlaps the text then that portion of the
    ' text will get inverted, then XOr it again in the background colour,
    ' if you use the same colour for the FillColor and ForeColor then the
    ' text will invert nicely, but you can get some funny effects if you
    ' use two different colours
    ' NOTE: treat 0% as a special case because it will show up as a 1
    ' pixel wide line which looks bad
    ' ALSO NOTE: I am using BF in the call to the Line method, which means to
    ' draw a filled box, although I only want to draw lines which are a
    ' single pixel thick, because with trial and error I have found that this
    ' gives me the lines where I expect them for the co-ords that I am passing
  End With 'pic
  If iPercent > 0 Then
    With pic
      .DrawMode = XORPEN
      ' XOr the pen
      pic.Line (iLeft, iTop)-((iRight / 100) * iPercent, iBottom), pic.FillColor, BF
      pic.Line (iLeft, iTop)-((iRight / 100) * iPercent, iBottom), pic.BackColor, BF
    End With 'pic
  End If
  ' add the 3D look - right, bottom, top, left
  With pic
    .DrawMode = COPYPEN
    pic.Line (iRight, iLineWidth)-(iRight, iBottom), WHITECOLOUR, BF
    pic.Line (iLineWidth, iBottom)-(iRight, iBottom), WHITECOLOUR, BF
    pic.Line (0, 0)-(iRight, 0), DGREYCOLOUR, BF
    pic.Line (0, 0)-(0, iBottom), DGREYCOLOUR, BF
    ' this line adds an additional grey border around the inside of the control to
    ' accentuate the 3D border - personal preference thing
    pic.Line (iLeft, iTop)-(iRight - iLineWidth, iBottom - iLineWidth), LGREYCOLOUR, B
  End With 'pic

End Sub

Private Sub HScroll_Change()

  picThumbnail.Left = -(HScroll.Value)

End Sub

Private Sub ImgListExtract(arrC As Variant, _
                           StrFileIn As String, _
                           ByVal lCStart As Long, _
                           ByVal lCEnd As Long, _
                           sImageData As String, _
                           sFolder As String, _
                           strFormName As String, _
                           strControlName As String, _
                           strControlProperty As String, _
                           nImageCount As Long)

  Dim I            As Long
  Dim PrevOffset   As Long
  Dim strListImage As String

  Dim strKey       As String
  Dim strTag       As String
  For I = lCStart To lCEnd
    If MatchString(CStr(arrC(I)), "BeginProperty ListImage") Then
      strListImage = Split(arrC(I))(1)
    End If
    sImageData = GetImageData(CStr(arrC(I)), StrFileIn, PrevOffset)
    'found an image so process it
    If LenB(sImageData) Then
      If MatchString(CStr(arrC(I + 1)), "Key ") Then
        strKey = GetStringValue(arrC(I + 1))
      End If
      If MatchString(CStr(arrC(I + 1)), "Object.Tag ") Then
        strTag = GetStringValue(arrC(I + 1))
      End If
      If MatchString(CStr(arrC(I + 2)), "Object.Tag ") Then
        strTag = GetStringValue(arrC(I + 2))
      End If
      If LenB(strKey) Then
        strKey = "Key(" & strKey & ")"
      End If
      If LenB(strTag) Then
        strKey = IIf(LenB(strKey), " ", vbNullString) & "Tag(" & strTag & ")"
      End If
      ProcessOneImage sImageData, sFolder, strFormName, strControlName & "_" & strListImage, strKey, strControlProperty, nImageCount
    End If
  Next I

End Sub

Private Sub imgThumbnail_Click(Index As Integer)

  SelectFrame Index
  CurImage = Index
  If LenB(PicData(CurImage).iName) Then
    cmdDelete.Enabled = True
    DisplayPicData CurImage
    picDisplay.Picture = LoadPicture(PicData(CurImage).ipath & "\" & PicData(CurImage).iName)
  End If

End Sub

Private Function IsFrxGraphicLine(ByVal sString As String, _
                                  strProp As String) As Boolean

  'Dim n As Long

  sString = Trim$(sString)
  If MatchString(sString, "Icon ") Then
    IsFrxGraphicLine = True
    strProp = "Icon"
   ElseIf MatchString(sString, "MouseIcon ") Then
    strProp = "MouseIcon"
    IsFrxGraphicLine = True
   ElseIf MatchString(sString, "DisabledPicture ") Then
    strProp = "MouseIcon"
    IsFrxGraphicLine = True
   ElseIf MatchString(sString, "DownPicture ") Then
    strProp = "MouseIcon"
    IsFrxGraphicLine = True
   ElseIf MatchString(sString, "Picture ") Then
    strProp = "Picture"
    IsFrxGraphicLine = True
   ElseIf MatchString(sString, "MaskPicture ") Then
    strProp = "MaskPicture"
    IsFrxGraphicLine = True
   ElseIf MatchString(sString, "TooBoxBitMap ") Then
    strProp = "ToolBoxBitMap"
    IsFrxGraphicLine = True
  End If

End Function

Private Sub lstSource_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

  Dim lIndex  As Long

  If Button = 0 Then ' Only if no button was pressed
    '
    '
    With lstSource
      ' Get selected item from list
      lIndex = SendLBMessage(.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((CLng(Y / Screen.TwipsPerPixelY) * 65536) + CLng(X / Screen.TwipsPerPixelX)))
      ' Show tip or clear last one
      If (lIndex >= 0) And (lIndex <= .ListCount) Then
        .ToolTipText = sSource(.ItemData(lIndex))
       Else
        .ToolTipText = ""
      End If
    End With
  End If

End Sub

Private Function MatchString(sExpression As String, _
                             sContaining As String, _
                             Optional ByVal bCaseSensitive As Boolean = False) As Boolean

  If bCaseSensitive Then
    MatchString = (Left$(sExpression, Len(sContaining)) = sContaining)
   Else
    MatchString = (Left$(UCase$(sExpression), Len(sContaining)) = UCase$(sContaining))
  End If

End Function

Private Sub mnuHelp_Click()

  MsgBox "Why?" & vbNewLine & _
       "1. Up/Downloading is easier if you don't include unneeded graphic files but what if you want to edit/borrow/reclaim an image?" & vbNewLine & _
       "2. Embedding Bitmaps make programs very large, converting to Jpeg and reloading into the program will save space." & vbNewLine & _
       "How?" & vbNewLine & _
       "1. Select the binary files you want (*.frx *.dox *.ctx *.dsx *pax). CommonDialog allows Multiselect" & vbNewLine & _
       "2. Click 'Start Extraction' The program extracts ALL recognised graphics(bmp, gif, ico, jpg, wmf (cur as ico, dib as bmp, jpeg & Tiff as jpg)) to the Destination folder (Default = source folder)." & vbNewLine & _
       "3. FileName format is 'FileName_ControlName_Property.FileExtention' to simplify reloading." & vbNewLine & _
       "   ControlName includes Index." & vbNewLine & _
       "   Imaglist Property is 'ListImage#'. Key and Tag values are included in brackets." & vbNewLine & _
       "   NOTE Jpgs that may have MS GDIPLUS.DLL JPEG EXPLOIT trigger warnings. RECOMMENDED you do not save AND delete the image from the control." & vbNewLine & _
       "4. Selecting images in the scroll viewer displays them in the main viewer." & vbNewLine & _
       "5. You can delete the file in the main viewer.", vbInformation, strBaseCaption

  'EXTENDED HELP
  'MS GDIPLUS.DLL JPEG EXPLOIT
  'becuase of the way this exploit works (overrun) this progra can't detect the file extention to use.
  'The file may contain the bug or be corrupted in some other way

End Sub

' [Borrowed code below...]
Private Sub picGuage_Paint()

  ' this event will only get fired if AutoRedraw = False

  If IsNumeric(picGuage.Tag) Then
    Guage picGuage, CInt(picGuage.Tag)
  End If

End Sub

Private Sub PositionThumbs()

  Dim thmb As Variant
  Dim pos  As Long

  For Each thmb In imgThumbnail
    If thmb.Visible Then
      thmb.Left = picThumbnail.Left + 460 * pos
      pos = pos + 1
    End If
  Next thmb
  If 460 * pos - 1 > HScroll.Width Then
    picThumbnail.Width = 460 * pos - 1
   Else
    picThumbnail.Width = HScroll.Width
  End If
  If picThumbnail.Width > HScroll.Width Then
    'increase the scoll width
    HScroll.Enabled = True
    With HScroll
      .Max = (picThumbnail.Width - .Width)
      .LargeChange = IIf(.Max < .Width, .Max, .Width)
      If .Value > .Max Then
        .Value = .Max
        HScroll_Change
      End If
    End With
   Else
    HScroll.Enabled = False
  End If
  picThumbnail.SetFocus

End Sub

Private Sub ProcessOneImage(sImageData As String, _
                            sFolder As String, _
                            strFormName As String, _
                            strControlName As String, _
                            strControlIndex As String, _
                            strControlProperty As String, _
                            nImageCount As Long)

  Dim sImageFile As String
  Dim strExt     As String

  strExt = GetImageExtention(sImageData)
  If LenB(strExt) Then
    sImageFile = AttachPath(strFormName & IIf(strFormName <> strControlName, "_" & strControlName, "_") & IIf(Len(strControlIndex), "(" & strControlIndex & ")", vbNullString) & "_" & strControlProperty & "." & strExt, sFolder)
    'safety back up(protects you from overwriting new images with older ones in the frx)
    If FileExist(sImageFile) Then
      File2BAK sImageFile
    End If
    SaveTheImage sImageFile, sImageData
    GeneratePicData sImageFile, sImageData, nImageCount
    DisplayPicData nImageCount
    GenerateNewThumb nImageCount, sImageFile
    nImageCount = nImageCount + 1
    ' Show in thumbnails...
    lblExtractVB.Caption = nImageCount & " image" & IIf(nImageCount = 1, vbNullString, "s") & " extracted"
    sImageData = ""
   Else
    If MsgBox("No valid file extention for the image data for " & strControlName & "." & strControlProperty & vbNewLine & _
          "The data may be corrupted OR If the image is viewable it may be a jpg file using the MS GDIPLUS.DLL JPEG EXPLOIT" & vbNewLine & _
          "It is RECOMMENDED that you delete the graphic from the control." & vbNewLine & _
          "Do you want to save it anyway (extention changed to JPGX)?", vbCritical + vbYesNo + vbDefaultButton2, "WARNING") = vbYes Then
      sImageFile = AttachPath(strFormName & IIf(strFormName <> strControlName, "_" & strControlName, "_") & IIf(Len(strControlIndex), "(" & strControlIndex & ")", vbNullString) & "_" & strControlProperty & ".JPGX", sFolder)
      'safety back up(protects you from overwriting new images with older ones in the frx)
      If MsgBox("In case the MS GDIPLUS.DLL JPEG EXPLOIT is present make sure you have XP Service Pack 2 before creating the file." & vbNewLine & _
          vbNewLine & _
          "OK to proceed?", vbCritical + vbOKCancel + vbDefaultButton2, "WARNING") = vbOK Then
        If FileExist(sImageFile) Then
          File2BAK sImageFile
        End If
        SaveTheImage sImageFile, sImageData
        GeneratePicData sImageFile, sImageData, nImageCount
        DisplayPicData nImageCount
        GenerateNewThumb nImageCount, sImageFile
        nImageCount = nImageCount + 1
        ' Show in thumbnails...
        lblExtractVB.Caption = nImageCount & " image" & IIf(nImageCount = 1, vbNullString, "s") & " extracted"
      End If
    End If
    'destroy the data
    sImageData = ""
  End If

End Sub

Private Sub ProgressBar(ByVal nPercent As Long)

  Guage picGuage, nPercent

End Sub

Private Sub SaveTheImage(ByVal strFile As String, _
                         ByVal strData As String)

  Dim FN As Long

  FN = FreeFile
  Open strFile For Binary Access Write Lock Write As FN
  Put #FN, 1, strData
  Close FN

End Sub

Private Sub SelectFrame(thumbNo As Variant)

  With shpSelector
    .Move imgThumbnail(thumbNo).Left, imgThumbnail(thumbNo).Top, imgThumbnail(thumbNo).Width, imgThumbnail(thumbNo).Height
    .ZOrder
    .Visible = True
  End With 'shpSelector

End Sub

Private Sub Set_OK_State()

  cmdStartExtraction.Enabled = (lstSource.ListCount > 0 And txtDestination.Text <> "")

End Sub

Private Function StripQuotes(ByVal sString As String) As String

  If Asc(Left$(sString, 1)) = 34 And Asc(Right$(sString, 1)) = 34 Then
    StripQuotes = Mid$(sString, 2, Len(sString) - 2)
   Else
    StripQuotes = sString
  End If

End Function


':)Code Fixer V2.9.2 (1/02/2005 4:13:46 PM) 23 + 907 = 930 Lines Thanks Ulli for inspiration and lots of code.

