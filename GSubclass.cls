VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright © 2015 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' GSubclass.bas
'
' Global Subclassing Interface to the DLL
'   - Dependancy: ISubclass.cls, MSubclass.bas
'   - Forwards calls to MSubclass
'
Option Explicit

Public Function SetSubclass(ByVal hWnd As Long, ByRef This As ISubclass, Optional ByVal dwRefData As Long) As Long
    SetSubclass = MSubClass.SetSubclass(hWnd, This, dwRefData)
End Function

Public Function RemoveSubclass(ByVal hWnd As Long, ByRef This As ISubclass) As Long
    RemoveSubclass = MSubClass.RemoveSubclass(hWnd, This)
End Function

Public Function DefSubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    DefSubclassProc = MSubClass.DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

