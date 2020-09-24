Attribute VB_Name = "Engine"

' written by
' Zaid Markabi , Arabic Syrian Student

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Engine
Global Tv3D As New TVEngine
Global Scene As New TVScene
Global Mesh As New TVMesh
Global Screen2DImmediate As New TVScreen2DImmediate

' Lighting
Global AddColorLighting1 As Integer
Global AddColorLighting2 As Integer
Global AddColorLighting3 As Integer
Global AddColorLighting4 As Integer
Global AddColorLightingR As Single
Global AddColorLightingG As Single
Global AddColorLightingB As Single

Global AddColorLightingA As Single


Sub Load_File_From_Res(ID As Integer, File_Name As String)
On Error Resume Next

Dim MyArry() As Byte
Dim MyFile As Long

If Dir$(File_Name) = "" Then
MyArry() = LoadResData(ID, "CUSTOM")
MyFile = FreeFile

Open File_Name For Binary Access Write As #1
Put #MyFile, , MyArry
Close #MyFile
End If
End Sub
