Attribute VB_Name = "modDeclares"
Public Declare Function SendMessage Lib "User32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63




Function CountItems(sField As String, sSep As String) As Integer
'Counts the number of items in a delimited string
'Pass the String that has the list, and the seperator
'iCount = sField, sSep
'sField = the list, sSep is the seperator to count by
    Dim bNotFound As Boolean
    Dim iPos As Integer, iCount As Integer
    
    Do Until bNotFound                                 'loop until we have counted all the items
        iPos = InStr(iPos + 1, sField, sSep)
        If iPos = 0 Then
            bNotFound = True                            'if we are done, then leave
        Else
            iCount = iCount + 1                         'increment the counter
        End If
    Loop
    CountItems = iCount                                 'return what we found
End Function

