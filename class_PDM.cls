VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_PDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'version 2025.0303

' Class module for product operations

' Public Class Product_Module

'




Private i  As Integer

Private bDict  As Object

Private rootPrd

Option Explicit



Sub init()

if

Set rootPrd = CATIA.ActiveDocument.Product



End Sub




Function whois2rv()
   '获取要读取或修改的  全局变量 Prd2wt 定义
    Dim infobox: infobox = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
        Select Case infobox
        Case 6  '===选择“是”,进行产品选择====
           On Error Resume Next
           Set whois2rv = selPrd()
           If Err.Number <> 0 Then
           Err.Clear
           Exit Function
           End If
        Case 7: Set whois2rv = CATIA.ActiveDocument.Product '===选择“否”====
        Case 2: Exit Function '===选择“取消”====
        End Select
        If Not whois2rv Is Nothing Then
           infobox = "你选择的产品是" & whois2rv.PartNumber & "是否继续"
          MsgBox infobox
        End If
           
End Function




















    ' Recursively process products and write information to Excel
    Public Sub recurPrd(oPrd, xlsht, oRowNb, LV)
        Dim bDict, i
        Arry2sht bomRowPrd(oPrd, LV), xlsht, oRowNb
        If oPrd.Products.Count > 0 Then ' I
            Set bDict = CreateObject("Scripting.Dictionary")
            For i = 1 To oPrd.Products.Count
                If Not bDict.Exists(oPrd.Products.Item(i).PartNumber) Then
                    bDict(oPrd.Products.Item(i).PartNumber) = 1
                    oRowNb = oRowNb + 1
                    recurPrd oPrd.Products.Item(i), xlsht, oRowNb, LV + 1
                End If
            Next
        End If
    End Sub

    ' Calculate the total mass of a product and its sub-products
    Public Function Assmass(oPrd)
        Dim i
        Dim total, children: total = 0
        Set children = oPrd.Products
        If oPrd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass = Assmass + Assmass(children.Item(i))
                total = total + children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value
            Next
            oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value
        End If
        Assmass = total
    End Function

    ' Generate a BOM row for a product
    Public Function bomRowPrd(oPrd, LV)
        Dim oDict, QTy
        QTy = 1
        If TypeOf oPrd.Parent Is Products Then ' If the product has a parent product
            Set oDict = Dictbros(oPrd.Parent.Parent)
            QTy = oDict(oPrd.PartNumber) ' Get the quantity of the product
        End If
        bomRowPrd = Array(LV, attDefault(oPrd), QTy, attUsp(oPrd))
    End Function

    ' Generate product information
    Public Function infoPrd(oPrd)
        Dim LV, QTy: LV = 0: QTy = 0
        infoPrd = Array(LV, attDefault(oPrd), QTy, attUsp(oPrd))
    End Function

    ' Get default attributes of a product
    Public Function attDefault(oPrd)
        With oPrd.ReferenceProduct
            Dim att_default(1 To 4)
            att_default(1) = .PartNumber
            att_default(2) = .Nomenclature
            att_default(3) = .Definition
            att_default(4) = oPrd.Name
            attDefault = att_default
        End With
    End Function

    ' Get user-defined attributes of a product
    Public Function attUsp(oPrd)
        Dim refprd: Set refprd = oPrd.ReferenceProduct
        Dim i, colls
        Dim att_usp(2 To 5)
        For i = 2 To 5
            Select Case i
                Case 2, 4, 5
                    Set colls = refprd.UserRefProperties
                    att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
                    Set colls = Nothing
                Case 3
                    On Error Resume Next
                    Dim oPrt
                    Set oPrt = refprd.Parent.Part
                    Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                    att_usp(i) = thisParameterValue(colls, att_usp_Names(i))
                    Set colls = Nothing
                    If Err.Number <> 0 Then
                        att_usp(i) = "N\A"
                        Err.Clear ' Handle non-part cases, cases without "cm", and cases without parameters
                    End If
                    On Error GoTo 0
            End Select
        Next
        attUsp = att_usp
    End Function

    ' Get a dictionary of sibling products
    Public Function Dictbros(oPrd)
        Dim oDict, i, pn: Set oDict = CreateObject("Scripting.Dictionary")
        For i = 1 To oPrd.Products.Count
            pn = oPrd.Products.Item(i).PartNumber
            If oDict.Exists(pn) Then
                oDict(pn) = oDict(pn) + 1
            Else
                oDict(pn) = 1
            End If
        Next
        Set Dictbros = oDict
    End Function

    ' Check if a collection has a specific parameter
    Public Function hasAtt(colls, parameterName)
        Dim propertyItem
        On Error Resume Next
        Set propertyItem = colls.Item(parameterName)
        If Err.Number <> 0 Then
            Err.Clear
            hasAtt = False
        Else
            hasAtt = True
        End If
        On Error GoTo 0
    End Function

    ' Get the value of a specific parameter in a collection
    Public Function thisParameterValue(colls, parameterName)
        On Error Resume Next
        thisParameterValue = colls.Item(parameterName).Value
        If Err.Number <> 0 Then
            Err.Clear
            thisParameterValue = "N\A"
        End If
        On Error GoTo 0
    End Function

l Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(Key)) > 0, True, False)
End Function

' 检查是否为字符串数组
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    
    IsStringAry = True
End Function

' 检查过滤器类型是否有效
Private Function IsFilterType(ByVal ary As Variant