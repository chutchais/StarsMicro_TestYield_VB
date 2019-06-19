VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Public Temperature As String
Public Total As Double
Public Pass As Double
Public Fail As Double
Public Yield As Double
Private objSwBins As New Collection

'Add on version 1.0.20 -- June 19,2019
Private objHwBins As New Collection

Public Property Get SwBinSums() As Collection
    Set SwBinSums = objSwBins
End Property

'Add on version 1.0.20 -- June 19,2019
Public Property Get HwBinSums() As Collection
    Set HwBinSums = objHwBins
End Property


Public Function getBin(key As String, col As Collection) As Object
  On Error GoTo errhandler
  For Each C In col
        If C.Name = key Then
            Set getBin = C
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set getBin = Nothing
End Function
