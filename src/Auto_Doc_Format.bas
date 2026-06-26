Sub Auto_Doc_Format()
    ' 关闭屏幕更新以提高代码运行速度
    Application.ScreenUpdating = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As Paragraph
    Dim txt As String
    Dim dotPos As Integer
    Dim i As Integer
    Dim isHeading As Boolean
    Dim char As String
    Dim numChars As String
    
    ' ===================================================
    ' 1. 全文格式：宋体，小四（12磅），固定行距20磅
    ' ===================================================
    With doc.Content
        .Font.NameFarEast = "宋体"
        .Font.NameAscii = "Times New Roman" ' 英文及数字推荐使用Times New Roman
        .Font.Size = 12                     ' Word中小四对应12磅
        .Font.Bold = False                  ' 取消可能存在的全局加粗
        .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
        .ParagraphFormat.LineSpacing = 20   ' 固定行距20磅
        .ParagraphFormat.SpaceBefore = 0    ' 取消段前距
        .ParagraphFormat.SpaceAfter = 0     ' 取消段后距
    End With
    
    ' ===================================================
    ' 2. 查找大标题并格式化：黑体四号加粗，段前段后0.5行
    ' ===================================================
    ' 用于匹配中文序号的字符库
    numChars = "一二三四五六七八九十百零"
    
    For Each para In doc.Paragraphs
        ' 获取段落文本并去除首尾空格
        txt = Trim(para.Range.text)
        isHeading = False
        
        ' 查找顿号“、”的位置
        dotPos = InStr(txt, "、")
        
        ' 校验：如果存在顿号，且顿号出现在开头部分（比如“一、”到“一百二十、”）
        If dotPos > 1 And dotPos <= 6 Then
            isHeading = True
            ' 检查顿号前的每一个字符是否都是中文数字
            For i = 1 To dotPos - 1
                char = Mid(txt, i, 1)
                If InStr(numChars, char) = 0 Then
                    isHeading = False
                    Exit For
                End If
            Next i
        End If
        
        ' 如果符合大标题特征，则修改该段落格式
        If isHeading Then
            ' 1. 设置字体格式
            With para.Range.Font
                .NameFarEast = "黑体"
                .NameAscii = "Arial"       ' 标题英文字体
                .Size = 14                 ' Word中四号对应14磅
                .Bold = True               ' 字体加粗
            End With
            
            ' 2. 设置段落格式（修改此处：全面使用 para.Range.ParagraphFormat 避免编译错误）
            With para.Range.ParagraphFormat
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = 20          ' 固定行距20磅
                .SpaceBefore = LinesToPoints(0.5) ' 段前0.5行
                .SpaceAfter = LinesToPoints(0.5)  ' 段后0.5行
            End With
        End If
    Next para
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    MsgBox "排版完成！" & vbCrLf & _
           "1. 全文已设为：宋体小四，20磅固定行距。" & vbCrLf & _
           "2. 大标题已设为：黑体四号加粗，段前段后0.5行。", vbInformation, "排版成功"
End Sub