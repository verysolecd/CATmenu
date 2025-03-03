VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cat_Macro_Menu_View 
   Caption         =   "UserForm1"
   ClientHeight    =   7994
   ClientLeft      =   119
   ClientTop       =   448
   ClientWidth     =   5418
   OleObjectBlob   =   "Cat_Macro_Menu_View.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cat_Macro_Menu_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cat_Macro_Menu_View
'   Caption = "UserForm1"
'   ClientHeight = 3500
'   ClientLeft = 120
'   ClientTop = 450
'   ClientWidth = 4560
'   OleObjectBlob   =   "Cat_Macro_Menu_View.frx":0000
'   StartUpPosition = 1    'CenterOwner
'End
'Attribute VB_Name = "Cat_Macro_Menu_View"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
' VBA CATIA V5 详细使用方法（版本）说明  by Kantoku
' Cat_Macro_Menu_View.frm
' 该文件用于实现菜单的UI界面

' 窗体边距
Private FrmMargin As Variant ' 上, 右, 下, 左 窗体边距调整值

' 窗体宽度调整值
Private Const ADJUST_F_W = 10
' 窗体高度调整值
Private Const ADJUST_F_H = 10

' 多页控件调整
Private Const ADJUST_M_W = 80 ' 多页控件宽度调整值
Private Const ADJUST_M_H = 90 ' 多页控件高度调整值
Private Const Tab_W = 45 ' Tab固定宽度
Private Const Tab_H = 20 ' TAB高度
Private Const Tab_frontsize = 10


' 按钮尺寸
Private Const BTN_W = 68 ' 按钮的固定宽度
Private Const BTN_H = 20 ' 单个按钮的高度

Private Const BTN_frontsize = 10



Private mBtns As Object ' 按钮事件集合

Option Explicit

' 设置窗体信息
Sub Set_FormInfo(ByVal InfoLst As Object, _
                 ByVal PageMap As Object, _
                 ByVal FormTitle As String, _
                 ByVal CloseType As Boolean)
                 
    ' 初始化窗体边距
    FrmMargin = Array(10, 10, 10, 0) ' 上, 右, 下, 左 窗体边距调整值
    
    ' 创建多页控件
    Dim MPgs As MultiPage
    Set MPgs = Me.Controls.Add("Forms.MultiPage.1", 1, True)
    
    
    
    Dim Pgs As Pages
     Set Pgs = MPgs.Pages
     Pgs.Clear
    
    Dim Key As Long, KeyStr As Variant
    Dim Pg As Page, PName As String
    Dim BtnInfos As Object, Info As Variant
    Dim Btns As Object: Set Btns = KCL.InitLst()
    Dim Btn As MSForms.CommandButton
    Dim BtnEvt As Button_Evt
    
    For Each KeyStr In InfoLst
        Key = CLng(KeyStr)
        If Not PageMap.Exists(Key) Then GoTo Continue
        PName = PageMap(Key)
        Set Pg = Get_Page(Pgs, PName)
        Set BtnInfos = InfoLst(KeyStr)
        For Each Info In BtnInfos
            Set Btn = Init_Button(Pg.Controls, Key, Info)
            Set BtnEvt = New Button_Evt
            Call BtnEvt.set_ButtonEvent(Btn, Info, Me, CloseType)
            Btns.Add BtnEvt
        Next
Continue:
    Next
    
    Set mBtns = Btns
    Call Set_MPage(MPgs)
    Call Set_Form(MPgs, FormTitle)
End Sub

' 设置窗体属性
Private Sub Set_Form(ByVal MPgs As MultiPage, ByVal Cap As String)
    With Me
        .Height = MPgs.Height + ADJUST_F_H
        .Width = MPgs.Width + ADJUST_F_W
        .Caption = Cap
    End With
End Sub

' 设置多页控件属性
Private Sub Set_MPage(ByVal MPgs As MultiPage)

    MPgs.Width = FrmMargin(1) + BTN_W + FrmMargin(3) + ADJUST_M_W
    With MPgs
        .TabFixedHeight = Tab_H  ' 标签高度（单位：磅）
        .TabFixedWidth = Tab_W ' 标签宽度
        .Font.Name = "Arial"
        .Font.Size = Tab_frontsize
        .MultiRow = True
        .TabOrientation = fmTabOrientationLeft
        
'        .Style = fmTabStyleButtons  ' 切换为按钮样式
    
     End With
    
    
    Dim Pg As Page
    Dim MaxBtnCnt As Long: MaxBtnCnt = 0
    Dim BtnCnt As Long
    For Each Pg In MPgs.Pages

    
    
        BtnCnt = Pg.Controls.Count
        MaxBtnCnt = IIf(BtnCnt > MaxBtnCnt, BtnCnt, MaxBtnCnt)
    Next
    
    MPgs.Height = FrmMargin(0) + (BTN_H * MaxBtnCnt) + FrmMargin(2) + ADJUST_M_H
    
    ' 设置多页控件背景颜色

End Sub

' 初始化按钮
Private Function Init_Button(ByVal Ctls As Controls, _
                             ByVal Idx As Long, _
                             ByVal BtnInfo As Variant) As MSForms.CommandButton
    Dim Btn As MSForms.CommandButton
    Set Btn = Ctls.Add("Forms.CommandButton.1", Idx, True)
    
    Dim Pty As Variant
    For Each Pty In BtnInfo.keys
        Call Try_SetProperty(Btn, Pty, BtnInfo.Item(Pty))
    Next
    
    With Btn
        .Top = FrmMargin(0) + (Ctls.Count - 1) * BTN_H
        .Left = FrmMargin(2)
        .Height = BTN_H
        .Width = BTN_W
        ' 设置按钮字体
        .Font.Name = "Arial"
        .Font.Size = BTN_frontsize
        ' 设置按钮背景颜色
       ' .BackColor = RGB(220, 220, 220)
    End With
    
    Set Init_Button = Btn
End Function

' 尝试设置控件属性
Private Sub Try_SetProperty(ByVal Ctrl As Object, _
                            ByVal PptyName As String, _
                            ByVal Value As Variant)
    On Error Resume Next
        Err.Number = 0
        
        Dim tmp As Variant
        tmp = CallByName(Ctrl, PptyName, VbGet)
        If Not Err.Number = 0 Then
            Debug.Print PptyName & ": 获取属性失败(" & Err.Number & ")"
            Exit Sub
        End If
        
        Select Case TypeName(tmp)
            Case "Empty"
                Exit Sub
            Case "Long"
                Value = CLng(Value)
            Case "Boolean"
                Value = CBool(Value)
            Case "Currency"
                Value = CCur(Value)
        End Select
        If Not Err.Number = 0 Then
            Debug.Print Value & ": 类型转换失败(" & Err.Number & ")"
            Exit Sub
        End If
        
        Call CallByName(Ctrl, PptyName, VbLet, Value)
        If Not Err.Number = 0 Then
            Debug.Print Value & ": 设置属性失败(" & Err.Number & ")"
            Exit Sub
        End If
    On Error GoTo 0
End Sub
' 获取页面 - 若不存在则创建
Private Function Get_Page(ByVal Pgs As Pages, ByVal Name As String) As Page
    Dim Pg As Page
    On Error Resume Next
        Set Pg = Pgs.Item(Name)
    On Error GoTo 0
    If Pg Is Nothing Then
        Set Pg = Pgs.Add(Name, Name, Pgs.Count)
    End If
    Set Get_Page = Pg
End Function


Private Sub TabStrip1_Change()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub selPrd_Click()

End Sub

Private Sub UserForm_Click()

End Sub


