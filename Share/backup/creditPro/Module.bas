Attribute VB_Name = "Module1"
Global Const DEFSOURCE = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source="
Public db As ADODB.Connection

Public Sub OpenDB()
    Set db = New ADODB.Connection
    db.Open DEFSOURCE & App.Path & "\db.MDB;"
    DEnv.Connection1 = DEFSOURCE & App.Path & "\db.MDB;"
End Sub
