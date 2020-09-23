VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Show Hand on Label"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Cursor is Hand #2- hover over me"
      Height          =   330
      Index           =   1
      Left            =   675
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Cursor is Hand #1- hover over me"
      Height          =   330
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   405
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' this example simply makes the hand cursor the form's primary cursor

' used to convert icons/bitmaps to stdPicture objects
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type
' used to load the current hand cursor theme
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Const IDC_HAND As Long = 32649

Private myHandCursor As StdPicture
Private myHand_handle As Long

Private Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture
' Convert an icon/bitmap handle to a Picture object

On Error GoTo ExitRoutine

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture

ExitRoutine:
End Function

Private Sub Form_Load()
    
    ' load the system's hand cursor if it exists
    myHand_handle = LoadCursor(0, IDC_HAND)
    If myHand_handle <> 0 Then
        ' use function to convert memory handle to stdPicture
        ' so we can apply it to the MouseIcon
        Set myHandCursor = HandleToPicture(myHand_handle, False)
    End If
    

    ' this is option #1
    ' Once we set the cursor, we'll never know if the user changed their
    ' mouse theme, but that happens rarely, option #2 you can find in the
    ' Label1 MouseMove function, which will always display the current themed
    ' hand icon regardless if it is changed while the app is running.

        If Not myHandCursor Is Nothing Then
            Label1(0).MouseIcon = myHandCursor
            Label1(0).MousePointer = vbCustom
        End If
    ' note that if you used replaced Label1 with Me, not only the form itself,
    ' but all controls on the form would show the hand cursor as long as the
    ' controls have their MousePointer=vbDefault.


End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 Then
        ' option #2
        ' We already got a handle to whatever icon will be used for the system's
        ' hand (link select) cursor. By placing the following in our control's
        ' mousemove event, it will always select the correct cursor even if it
        ' is changed while this app is running. Try it.
        SetCursor myHand_handle
    End If
End Sub
