Option Compare Database
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" (ByVal lpStr1 As LongPtr, ByVal lpStr2 As Any) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
    Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32.dll" (ByVal lpStr1 As Any, ByVal lpStr2 As Any) As Long
#End If

Private Const CF_TEXT As Long = 1&
Private Const GMEM_MOVEABLE As Long = 2

Private Sub Beispiel()
    Call StringToClipboard("Hallo ...")
End Sub

Public Sub StringToClipboard(strText As String) ' $00:00001
    #If VBA7 Then
        Dim lngIdentifier As LongPtr, lngPointer As LongPtr
    #Else
        Dim lngIdentifier As Long, lngPointer As Long
    #End If
    
    lngIdentifier = GlobalAlloc(GMEM_MOVEABLE, Len(strText) + 1)
    lngPointer = GlobalLock(lngIdentifier)
    Call lstrcpy(ByVal lngPointer, strText)
    Call GlobalUnlock(lngIdentifier)
    Call OpenClipboard(0&)
    Call EmptyClipboard
    Call SetClipboardData(CF_TEXT, lngIdentifier)
    Call CloseClipboard
    Call GlobalFree(lngIdentifier)
End Sub