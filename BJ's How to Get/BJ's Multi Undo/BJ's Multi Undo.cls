VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -- Current Key
Private strCurrentDirectory As String
Public SelStart As Long  'start position in text box
Public TextLen As Long
Public Text As String
'
' -----------------------------------------------------------------------
'
'
Private Sub CdExt(ByVal strDirectory As String, ByVal lngSecurity As Long)

Dim lngNewKeyHandle As Long
Dim lngPos As Long
  On Error GoTo CdExt_Err
  strDirectory = Trim(UCase(strDirectory))
  Select Case strDirectory
    Case "\"
      If strCurrentDirectory <> "" Then _
        TestAPIError RegCloseKey(lngKeyHandle)
      strCurrentDirectory = ""
      lngKeyHandle = 0
    Case ".."
      If strCurrentDirectory = "" Then ClassError 1
      lngPos = 1
      Do While InStr(lngPos, strCurrentDirectory, "\") > 0
        lngPos = InStr(lngPos, strCurrentDirectory, "\") + 1
      Loop
      CdExt Mid(strCurrentDirectory, 1, lngPos - 1), lngSecurity
    Case Else ' CdExt <path>
      
      If Right(strDirectory, 1) = "\" And Len(strDirectory) > 1 Then _
        strDirectory = Mid(strDirectory, 1, Len(strDirectory) - 1)
      
      If strCurrentDirectory = "" And Left(strDirectory, 1) <> "\" Then _
        strDirectory = "\" & strDirectory
      If Left(strDirectory, 1) = "\" Then
        If strCurrentDirectory <> "" Then TestAPIError RegCloseKey(lngKeyHandle)
        Select Case strRoot(strDirectory)
          Case "HKEY_CLASSES_ROOT"
            lngKeyHandle = HKEY_CLASSES_ROOT
          Case "HKEY_CURRENT_USER"
            lngKeyHandle = HKEY_CURRENT_USER
          Case "HKEY_LOCAL_MACHINE"
            lngKeyHandle = HKEY_LOCAL_MACHINE
          Case "HKEY_USERS"
            lngKeyHandle = HKEY_USERS
          Case "HKEY_DYN_DATA"
            lngKeyHandle = HKEY_DYN_DATA
          Case "HKEY_CURRENT_CONFIG"
            lngKeyHandle = HKEY_CURRENT_CONFIG
          Case Else
            ClassError 2
        End Select
        strCurrentDirectory = strDirectory
        TestAPIError _
          RegOpenKeyEx _
          (lngKeyHandle, strPath(strCurrentDirectory), 0, lngSecurity, lngNewKeyHandle)
        lngKeyHandle = lngNewKeyHandle
      Else ' Percorso dalla posizione corrente
        If strCurrentDirectory = "" Then ClassError 1 'i changed this
        strCurrentDirectory = strCurrentDirectory & "\" & strDirectory
        TestAPIError _
          RegOpenKeyEx _
          (lngKeyHandle, strDirectory, 0, lngSecurity, lngNewKeyHandle)
        lngKeyHandle = lngNewKeyHandle
      End If
  End Select
  'Debug.Print strCurrentDirectory, lngKeyHandle
  Exit Sub

CdExt_Err:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
