VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HIEN_THI_TUNG_CELL 
   Caption         =   "UserForm1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20340
   OleObjectBlob   =   "HIEN_THI_TUNG_CELL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HIEN_THI_TUNG_CELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#End If

Dim cot As Integer
Dim hang As Integer
Dim cot_Max As Integer
Dim hang_Max As Integer


Dim cot_Min As Integer
Dim hang_Min As Integer

Dim cot_Update As Integer
Dim hang_Update As Integer

Dim keyError() As String


Sub CapNhat_Click()
    Cells(cot_Update, hang_Update) = NoiDung.Text
End Sub




Sub CommandButton1_Click()

    cot_Min = CInt(CotBatDau.Text)
    cot_Max = CInt(CotKetThuc.Text)

    hang_Min = CInt(HangBatDau.Text)
    hang_Max = CInt(HangKetThuc.Text)
    
    'Get
    'Debug.Print (cot)
    NoiDung.Text = Cells(hang, cot)
    ActiveSheet.Cells(hang, cot).Select
    cot_Update = hang
    hang_Update = cot
    
    'Check
    Dim i As Integer
    Dim strCheck As String
    strCheck = NoiDung.Text
    For i = 0 To ArrayLen(keyError) - 1
        If InStr(UCase(strCheck), UCase(keyError(i))) > 0 Then
            MsgBoxUni ("Have Error: " & UniConvert(keyError(i), "Telex"))
            'MsgBox ("Len:" & Len(keyError(i)))
        End If
        'Cells(2, 2) = UCase(keyError(i))
    Next
    
    
    'Next
    cot = cot + 1
    While (Cells(hang, cot) = "")
        cot = cot + 1
        If cot >= cot_Max Then
            cot = cot_Min
            hang = hang + 1
             If hang >= hang_Max Then
                hang = hang_Min
                Dim Result As Integer
                Result = MsgBox("Do you want to continue?", vbYesNo)
                If Result = vbYes Then
                Else
                    Unload Me
                    GoTo END_FIND
                End If
            End If
        End If
    Wend
    
END_FIND:
    
    Debug.Print ("-------------------")
    Debug.Print ("Hang: " & hang)
    Debug.Print ("Cot: " & cot)
End Sub


Sub UserForm_Initialize()
    Debug.Print ("Init Form")
    
    CotBatDau.Text = "1"
    CotKetThuc.Text = "10"
    HangBatDau.Text = "1"
    HangKetThuc.Text = "248"
    
    cot_Min = CInt(CotBatDau.Text)
    cot_Max = CInt(CotKetThuc.Text)

    hang_Min = CInt(HangBatDau.Text)
    hang_Max = CInt(HangKetThuc.Text)
    
    cot = cot_Min
    hang = hang_Min
    
    'Check
    Dim strFilePath As String
    strFilePath = "E:\01_Data_Source\TOOL\Excel_Show_1Cell_moreEnter\ErorListString.TXT"
    Dim strData As String
    strData = ReadFile(strFilePath)
    'ReDim keyError(0 To 999)
    keyError = Split(strData, vbCrLf)
    'Dim i As Integer
    For i = 0 To ArrayLen(keyError) - 1
        keyError(i) = Replace(keyError(i), " ", "")
        'MsgBox (keyError(i))
    Next
End Sub


Public Function ArrayLen(arr() As String) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function UTF8to16(str As String) As String
    Dim position As Long, strConvert As String, codeReplace As Integer, strOut As String

    strOut = str
    position = InStr(strOut, Chr(195))

    If position > 0 Then
        Do Until position = 0
            strConvert = Mid(strOut, position, 2)
            codeReplace = Asc(Right(strConvert, 1))
            If codeReplace < 255 Then
                strOut = Replace(strOut, strConvert, Chr(codeReplace + 64))
            Else
                strOut = Replace(strOut, strConvert, Chr(34))
            End If
            position = InStr(strOut, Chr(195))
        Loop
    End If

    UTF8to16 = strOut
End Function


Public Function ReadFile(path As String, Optional CharSet As String = "utf-8")
  Static obj As Object
  If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
  obj.CharSet = CharSet
  obj.Open
  obj.LoadFromFile path
  ReadFile = obj.ReadText()
  obj.Close
End Function

Public Sub WriteFile(path As String, Text As String, Optional CharSet As String = "utf-8")
  Static obj As Object
  If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
  obj.CharSet = CharSet
  obj.Open
  obj.WriteText Text
  obj.SaveToFile path
  obj.Close
End Sub


Function UniConvert(Text As String, InputMethod As String) As String
  Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
  UniConvert = Text
  VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
      "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
      "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
      "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
      "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
  Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
      "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
      "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
      "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
      "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
  CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
      ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
      ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
      ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
      ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
      ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
      ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
      ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
  Select Case InputMethod
    Case Is = "VNI": Temp = VNI_Type
    Case Is = "Telex": Temp = Telex_Type
  End Select
  For i = 0 To UBound(CharCode)
    UniConvert = Replace(UniConvert, Temp(i), CharCode(i))
    UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
  Next i
End Function

Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant = vbNullString) As VbMsgBoxResult
   'Function MsgBoxUni(ByVal PromptUni As Variant, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal TitleUni As Variant, Optional HelpFile, Optional Context) As VbMsgBoxResult
   'BStrMsg,BStrTitle : La chuoi Unicode
   Dim BStrMsg, BStrTitle
   'Hàm StrConv Chuyen chuoi ve ma Unicode
   BStrMsg = StrConv(PromptUni, vbUnicode)
   BStrTitle = StrConv(TitleUni, vbUnicode)
   MsgBoxUni = MessageBoxW(GetActiveWindow, BStrMsg, BStrTitle, Buttons)
End Function
