Attribute VB_Name = "modGeneral"
Option Explicit
Public OCX() As Byte
Public bGPS As Boolean
Public sStatState As String
Public sStatArea As String
Public sStatCountry As String
Public sStatRegion As String
Public sStatCounty As String
Public PlayRegAnimation As Boolean
Public PlayAnimation As Boolean
Public AnimationLink As String
Public Animation As Boolean
Public sMapPicture As String
Public sFlagPicture As String
Public PictureName As String
Public scntName As String
Public iMinCount As Integer
Public sFrmName As String
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
        'Constants
'Const LB_FINDSTRINGEXACT = &H1A2    'To locate exact match

'Declares
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As String) As Long


Public Function FindStringinListControl(ListControl As Object, _
  ByVal SearchText As String) As Long

  '**************************************
  'Input:
  'ListControl: List or ComboBox Object
  'SearchText: String to Search For

  'Returns: ListIndex of Item if found
  'or -1 if not found
  '***************************************
  
  Dim lHwnd As Long
  Dim lMsg As Long

  'On Error Resume Next
  lHwnd = ListControl.hWnd

  If TypeOf ListControl Is ListBox Then
    lMsg = LB_FINDSTRINGEXACT
  ElseIf TypeOf ListControl Is ComboBox Then
    lMsg = CB_FINDSTRINGEXACT
  Else
    FindStringinListControl = -1
    Exit Function
  End If
  FindStringinListControl = SendMessageAsString(lHwnd, lMsg, -1, SearchText)
End Function

Public Function FileExists(FileName As String) As Boolean
  FileExists = (Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function SystemDirectory() As String
  Dim RSTR As String
  Dim RLEN As Long

  RSTR = String(255, 0)
  RLEN = GetSystemDirectory(RSTR, Len(RSTR))
  If RLEN < Len(RSTR) Then
    RSTR = Left(RSTR, RLEN)
    If Right(RSTR, 1) = "\" Then
      SystemDirectory = Left(RSTR, Len(RSTR) - 1)
    Else
      SystemDirectory = RSTR
    End If
  Else
    SystemDirectory = ""
  End If
End Function
