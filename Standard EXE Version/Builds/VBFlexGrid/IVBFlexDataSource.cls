VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVBFlexDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetFieldCount() As Long
Attribute GetFieldCount.VB_Description = "Interface function to define the number of fields."
End Function

Public Function GetRecordCount() As Long
Attribute GetRecordCount.VB_Description = "Interface function to define the number of records."
End Function

Public Function GetFieldName(ByVal Field As Long) As String
Attribute GetFieldName.VB_Description = "Interface function to define the field names."
End Function

Public Function GetData(ByVal Field As Long, ByVal Record As Long) As String
Attribute GetData.VB_Description = "Interface function for providing data to the consumer."
End Function

Public Sub SetData(ByVal Field As Long, ByVal Record As Long, ByVal NewData As String)
Attribute SetData.VB_Description = "Interface method for updating the new data information supplied by the consumer."
End Sub
