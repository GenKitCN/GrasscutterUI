VERSION 5.00
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmOperate 
   BackColor       =   &H80000005&
   Caption         =   "Grasscutter UI"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOperate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7920
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "搜"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      Height          =   420
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "生成"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Tag             =   "TL"
      Top             =   3480
      Width           =   1575
   End
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   6240
      Top             =   3480
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Tag             =   "HW"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   2680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2190
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OperateType As String
Public HandbookKey As String

Public Sub TriggerEnemy()
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & " - 生成敌怪"
    cmdAction.Caption = "生成"
    OperateType = "spawn"
    If Handbook.Exists("Monsters") Then
        HandbookKey = "Monsters"
    ElseIf Handbook.Exists("讨伐列表") Then
        HandbookKey = "讨伐列表"
    Else
        HandbookKey = InputBox("手册不支持，请手动输入敌怪类名")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Label1.Visible = True
    Label1.Caption = "[等级]:"
    Text1.Visible = True
    Label2.Visible = True
    Label2.Caption = "(等级为选填)"
    Me.Show
End Sub
Public Sub TriggerChar()
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & " - 给予角色"
    cmdAction.Caption = "给予"
    OperateType = "givechar"
    If Handbook.Exists("Characters") Then
        HandbookKey = "Characters"
    ElseIf Handbook.Exists("角色") Then
        HandbookKey = "角色"
    Else
        HandbookKey = InputBox("手册不支持，请手动输入角色类名")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Label1.Visible = True
    Label1.Caption = "[Amount]:"
    Text1.Visible = True
    Me.Show
End Sub
Public Sub TriggerScene()
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & " - 场景传送"
    cmdAction.Caption = "传送"
    OperateType = "changescene"
    If Handbook.Exists("Scenes") Then
        HandbookKey = "Scenes"
    ElseIf Handbook.Exists("Scene") Then
        HandbookKey = "Scene"
    Else
        HandbookKey = InputBox("手册不支持，请手动输入场景类名")
    End If
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Me.Show
End Sub

Public Sub TriggerAccount()
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & "  - 注册或删除账号"
    cmdAction.Caption = "执行"
    OperateType = "account"
    lst.AddItem "注册账号"
    lst.AddItem "删除账号"
    Label1.Visible = True
    Label1.Caption = "邮箱:"
    Text1.Visible = True
    Label2.Visible = True
    Label2.Caption = "(随意填写即可)"
    Text2.Visible = False
    Label3.Visible = True
    Label3.Caption = "UID:"
    Text3.Visible = True
    Text2.Text = "10001"
    Me.Show
End Sub

Public Sub TriggerStats()
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & "  - 角色数值修改"
    cmdAction.Caption = "执行"
    OperateType = "setstats"
    lst.AddItem "生命值 (ID:hp)"
    lst.AddItem "防御力 (ID:def)"
    lst.AddItem "攻击力 (ID:atk)"
    lst.AddItem "元素精通 (ID:em)"
    lst.AddItem "元素充能效率 (ID:er)"
    lst.AddItem "暴击率 (ID:crate)"
    lst.AddItem "暴击伤害 (ID:cdmg)"
    lst.AddItem "火元素伤害加成 (ID:epyro)"
    lst.AddItem "冰元素伤害加成 (ID:ecyro)"
    lst.AddItem "水元素伤害加成 (ID:ehydro)"
    lst.AddItem "岩元素伤害加成 (ID:egeo)"
    lst.AddItem "草元素伤害加成 (ID:edend)"
    lst.AddItem "雷元素伤害加成 (ID:eelec)"
    lst.AddItem "物理伤害加成 (ID:ephys)"
    Label1.Visible = True
    Label1.Caption = "数值:"
    Text1.Visible = True
    Me.Show
End Sub

Public Sub TriggerDrop()
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & "  - 生成掉落物"
    cmdAction.Caption = "给予"
    OperateType = "drop"
    Combo1.Clear
    Combo1.AddItem "武器"
    Combo1.AddItem "材料"
    Combo1.AddItem "圣遗物"
    Combo1.Text = "武器"
    HandbookKey = "武器"
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Combo1.Visible = True
    Label1.Visible = True
    Label1.Caption = "分类:"
    Text1.Visible = False
    Label2.Visible = True
    Label2.Caption = "数量:"
    Text2.Visible = True
    Me.Show
End Sub
Public Sub TriggerGive()
    If HandbookLoaded <> True Then MsgBox "手册正在加载中，请稍等几秒钟。", vbCritical: Exit Sub
    Me.Caption = "生草机 " & App.Major & "." & App.Minor & "." & App.Revision & "  - 给予物品"
    cmdAction.Caption = "给予"
    OperateType = "give"
    Combo1.Clear
    Combo1.AddItem "武器"
    Combo1.AddItem "材料"
    Combo1.AddItem "圣遗物"
    Combo1.Text = "武器"
    HandbookKey = "武器"
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
    Combo1.Visible = True
    Label1.Visible = True
    Label1.Caption = "分类:"
    Text1.Visible = False
    Label2.Visible = True
    Label2.Caption = "数量:"
    Text2.Visible = True
    Label3.Visible = True
    Label3.Caption = "[等级]:"
    Text3.Visible = True
    Label4.Visible = True
    Label4.Caption = "(等级为选填)"
    Me.Show
End Sub


Public Sub Search()
    lst.Clear
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        If InStr(Handbook(HandbookKey)(EachItem), txtSearch.Text) Then lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
End Sub

Private Sub cmdAction_Click()
If lst.Text = "" Then Exit Sub
Dim ItemID As String
Select Case OperateType
Case "spawn"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!spawn " & ItemID & " " & Text1.Text & vbCrLf & "已经复制到剪贴板。"
    Clipboard.SetText "!spawn " & ItemID & " " & Text1.Text
Case "give"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    If Text2.Text = "" Then Text2.Text = "1"
    frmMain.Server.SendLine "!give " & frmMain.txtUID.Text & " " & ItemID & " " & Text2.Text & " " & Text3.Text
    MsgBox "将 " & Text2.Text & " 个 " & lst.Text & " 给予 UID 为 " & frmMain.txtUID.Text & " 的玩家。"
Case "givechar"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    If Text1.Text = "" Then Text1.Text = "1"
    frmMain.Server.SendLine "!givechar " & frmMain.txtUID.Text & " " & ItemID & " " & Text1.Text
    MsgBox "将 " & lst.Text & " 给予 UID 为 " & frmMain.txtUID.Text & " 的玩家。"
Case "account"
    If Text1.Text = "" Then Exit Sub
    If lst.Text = "注册账号" Then
        If Text2.Text = "" Then Exit Sub
        If CInt(Text2.Text) <> Text2.Text Then Exit Sub
        frmMain.Server.SendLine "!account create " & Text1.Text & " " & Text2.Text
        frmMain.txtUID.Text = Text2.Text
        WriteIni "GCUI", "PlayerUID", Text2.Text, App.Path & "\Config.ini"
        MsgBox "注册了账号，邮箱地址为 " & Text1.Text & "，UID 为 " & Text2.Text & "。"
    ElseIf lst.Text = "删除账号" Then
        frmMain.Server.SendLine "!account delete " & Text1.Text & " " & Text2.Text
        frmMain.txtUID.Text = Text2.Text
        WriteIni "GCUI", "PlayerUID", Text2.Text, App.Path & "\Config.ini"
        MsgBox "删除了账号，邮箱地址为 " & Text1.Text & "。"
    Else
        Exit Sub
    End If
Case "changescene"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!changescene " & ItemID & vbCrLf & "已经复制到剪贴板。"
    Clipboard.SetText "!changescene " & ItemID
Case "setstats"
    If Text1.Text = "" Then Exit Sub
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!setstats " & ItemID & " " & Text1.Text & vbCrLf & "已经复制到剪贴板。"
    Clipboard.SetText "!setstats " & ItemID & " " & Text1.Text
Case "drop"
    ItemID = Mid(lst.Text, InStr(1, lst.Text, " (ID:") + 5, InStr(1, lst.Text, ")") - (InStr(1, lst.Text, " (ID:") + 5))
    MsgBox "!drop " & ItemID & " " & Text1.Text & vbCrLf & "已经复制到剪贴板。"
    Clipboard.SetText "!drop " & ItemID & " " & Text1.Text
Case Else
    MsgBox "出现了未知错误，请把这个窗口截图发给作者。" & vbCrLf & vbCrLf & "OperateType: " & OperateType & vbCrLf & Me.Caption
    End
End Select
Unload Me
End Sub

Private Sub cmdSearch_Click()
    Search
End Sub


Private Sub Combo1_Click()
    If Handbook.Exists(Combo1.Text) Then
        HandbookKey = Combo1.Text
    Else
        HandbookKey = InputBox("手册不支持，请手动输入敌怪类名")
        Combo1.Text = HandbookKey
    End If
    txtSearch.Text = ""
    lst.Clear
    Dim EachItem As Variant
    For Each EachItem In Handbook(HandbookKey).Keys
        lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
    Next
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtSearch.Text = "" Then
            lst.Clear
            Dim EachItem As Variant
            For Each EachItem In Handbook(HandbookKey).Keys
                lst.AddItem Handbook(HandbookKey)(EachItem) & " (ID:" & EachItem & ")"
            Next
        Else
            Search
        End If
        End If
    End Sub
