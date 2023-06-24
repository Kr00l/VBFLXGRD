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

Public Function GetFlags() As Long
Attribute GetFlags.VB_Description = "Interface function to get options from the data provider."
End Function

Public Function GetToolTipText(ByVal Field As Long, ByVal Record As Long) As String
Attribute GetToolTipText.VB_Description = "Interface function for providing tool tip text to the consumer."
End Function

Public Sub SetToolTipText(ByVal Field As Long, ByVal Record As Long, ByVal NewValue As String)
Attribute SetToolTipText.VB_Description = "Interface method for updating the new tool tip text supplied by the consumer."
End Sub

Public Function GetChecked(ByVal Field As Long, ByVal Record As Long) As Integer
Attribute GetChecked.VB_Description = "Interface function for providing check mark to the consumer."
End Function

Public Sub SetChecked(ByVal Field As Long, ByVal Record As Long, ByVal NewValue As Integer)
Attribute SetChecked.VB_Description = "Interface method for updating the new check mark supplied by the consumer."
End Sub
