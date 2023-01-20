VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVBFlexDataSource2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function NoFieldNames() As Boolean
Attribute NoFieldNames.VB_Description = "Interface function to indicate whether or not the data source has field names."
End Function

Public Function GetToolTipText(ByVal Field As Long, ByVal Record As Long) As String
Attribute GetToolTipText.VB_Description = "Interface function for providing tool tip text to the consumer."
End Function

Public Sub SetToolTipText(ByVal Field As Long, ByVal Record As Long, ByVal NewValue As String)
Attribute SetToolTipText.VB_Description = "Interface method for updating the new tool tip text supplied by the consumer."
End Sub
