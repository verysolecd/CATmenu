VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Button_Evt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute mBtn.VB_VarHelpID = -1

Option Explicit

Private WithEvents mBtn As MSForms.CommandButton
Private mInfo As Object
Private mParentForm As MSForms.UserForm
Private mCloseType As Boolean

Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签

' 设置事件
Sub set_ButtonEvent(ByVal Btn As MSForms.CommandButton, _
              ByVal BtnInfo As Object, _
              ByVal ParentForm As MSForms.UserForm, _
              ByVal CloseType As Boolean)
    Set mBtn = Btn
    Set mInfo = BtnInfo
    Set mParentForm = ParentForm
    mCloseType = CloseType
End Sub
' 按钮点击事件处理
Private Sub mBtn_Click()
    If Not mCloseType Then Unload mParentForm    
    Dim ary: ary = Array()
    Dim ss: Set ss = CATIA.SystemService    
    Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                          catScriptLibraryTypeVBAProject, _
                          mInfo(TAG_MDLNAME), _
                          mInfo(TAG_ENTRYPNT), _
                          Array())
End Sub




















' 示例方法：向工作表写入数据
Public Sub WriteDataToSheet(data As Variant, row As Long, col As Long)
    If Not m_Worksheet Is Nothing Then
        m_Worksheet.Cells(row, col).Value = data
    End If
End Sub











' 示例方法：从工作表读