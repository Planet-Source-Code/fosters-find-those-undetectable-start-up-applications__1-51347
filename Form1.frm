VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry quick check..."
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3615
      TabIndex        =   1
      Top             =   3840
      Width           =   2715
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6482
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15987699
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Key"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Company Name"
         Object.Width           =   38100
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GetAllValuesInAKey(ByVal sKeyDesc As String, ByVal HKEY As Long, ByVal sIN As String)
Dim Values As Variant
Dim KeyLoop As Integer
Dim RegPath As String
Dim sLine As String
Dim VerInfo As CFileVersionInfo
Dim sWinDir As String
Dim sTemp As String

If DirExist("c:\winnt") Then
    sWinDir = "c:\winnt\system32\"
Else
    sWinDir = "c:\windows\system32\"
End If

        
Values = GetAllValues(HKEY, sIN)

If VarType(Values) = vbArray + vbVariant Then
If Not IsArrayEmpty(Values) Then
For KeyLoop = 0 To UBound(Values)
    LV.ListItems.Add , , sKeyDesc
    LV.ListItems(LV.ListItems.Count).SubItems(1) = Values(KeyLoop, 0)
    Select Case Values(KeyLoop, 1)
    Case REG_DWORD
        LV.ListItems(LV.ListItems.Count).SubItems(2) = GetSettingLong(HKEY, sIN, _
        CStr(Values(KeyLoop, 0)))
    Case REG_BINARY
        LV.ListItems(LV.ListItems.Count).SubItems(2) = GetSettingByte(HKEY, sIN, _
        Hex$(Values(KeyLoop, 0)))(0)
    Case REG_SZ
        LV.ListItems(LV.ListItems.Count).SubItems(2) = GetSettingString(HKEY, sIN, _
        CStr(Values(KeyLoop, 0)))
    End Select
    sTemp = LV.ListItems(LV.ListItems.Count).SubItems(2)
    If Len(LV.ListItems(LV.ListItems.Count).SubItems(2)) > 0 Then
        Set VerInfo = New CFileVersionInfo
        If InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".exe") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".exe") + 3)
        ElseIf InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".com") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".com") + 3)
        ElseIf InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".bat") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".bat") + 3)
        ElseIf InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".cmd") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".cmd") + 3)
        ElseIf InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".dll") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".dll") + 3)
        ElseIf InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".scr") > 0 Then
            sTemp = Left(LV.ListItems(LV.ListItems.Count).SubItems(2), InStr(LCase(LV.ListItems(LV.ListItems.Count).SubItems(2)), ".scr") + 3)
        End If
        If InStr(LV.ListItems(LV.ListItems.Count).SubItems(2), "\") = 0 Then
            VerInfo.FullPathName = sWinDir & sTemp
        Else
            VerInfo.FullPathName = sTemp
        End If
        If VerInfo.Available Then
            LV.ListItems(LV.ListItems.Count).SubItems(3) = VerInfo.PredefinedValue(viCompanyName)
        End If
    End If
Next
End If
End If

End Sub

Private Sub Command1_Click()
    Unload Me
    End
    
End Sub

Private Sub Form_Load()
LV.ColumnHeaders(1).Width = ((LV.Width - 300) / 100) * 18
LV.ColumnHeaders(2).Width = ((LV.Width - 300) / 100) * 21
LV.ColumnHeaders(3).Width = ((LV.Width - 300) / 100) * 41
LV.ColumnHeaders(4).Width = ((LV.Width - 300) / 100) * 20

GetAllValuesInAKey "HKLM/Run", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetAllValuesInAKey "HKLM/RunOnce", HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetAllValuesInAKey "HKCU/Run", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run"
GetAllValuesInAKey "HKCU/RunOnce", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce"

End Sub
