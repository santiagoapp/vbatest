VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000

'Windows API calls to handle windows
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Public Sub RemoveCloseButton(objForm As Object)
    Dim lngStyle As LongPtr
    Dim lngHWnd As LongPtr

    Dim lpClassName As String
    lpClassName = vbNullString
    If Val(Application.Version) >= 9 Then
       lpClassName = "ThunderDFrame"
    Else
       lpClassName = "ThunderXFrame"
    End If

    lngHWnd = FindWindow(lpClassName, objForm.Caption)
    lngStyle = GetWindowLongPtr(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLongPtr lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If
End Sub

Public Function fillListOrComboBox(ByVal arr As Variant, ByVal ListOrComboBox As Object)
    ListOrComboBox.Clear
    If VarType(arr) <> vbEmpty Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            ListOrComboBox.AddItem arr(0, j)
            For i = LBound(arr, 1) To UBound(arr, 1)
                If VarType(arr(i, j)) = vbNull Then
                    ListOrComboBox.List(j, i) = ""
                Else
                    ListOrComboBox.List(j, i) = arr(i, j)
                End If
            Next i
        Next j
    End If
End Function


Public Function Calendar() As String
    
    Dim CalendarModal As Object
    Set CalendarModal = UserForms.Add("DatePicker")
    CalendarModal.show
    Calendar = CalendarModal.btnResponse
    Unload CalendarModal
    
End Function
