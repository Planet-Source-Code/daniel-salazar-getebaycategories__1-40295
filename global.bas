Attribute VB_Name = "Module1"
'This code fires off request to HTML page
'returns html page in text string
'Makes use of the XMLHTTPRequest object contained in msxml.dll.
'Check off Reference to MSXML Version 2.6
' I am using latest Version of IE5
' should also work with IE5.0 MSXML ver 2.0

Function GetHTTPFile(URL As String) As String
    Dim oHttp As Object
    'make use of the XMLHTTPRequest object contained in msxml.dll
    Set oHttp = CreateObject("Microsoft.XMLHTTP")
    
    'oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'oHttp.setRequestHeader "Content-Type", "text/xml"
    'oHttp.setRequestHeader "Content-Type", "multipart/form-data"
      
    oHttp.Open "GET", URL, False
    oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oHttp.send
  
    'check the feedback from the Net File
    Debug.Print "Ready State =" & oHttp.readyState
    'normal state =4
    Debug.Print "Status =" & oHttp.Status
    'normal status = 200
    Debug.Print "Status Text =" & oHttp.statusText
    Debug.Print oHttp.getAllResponseHeaders()
    'Debug.Print "Response Body =" & oHttp.responseBody
    Debug.Print "Response Body =" & StrConv(oHttp.responseBody, vbUnicode)
    'Debug.Print "Response Text =" & oHttp.responseText

    GetHTTPFile = StrConv(oHttp.responseBody, vbUnicode)
    Set oHttp = Nothing
End Function

Function CategoriesOnly(html As String) As String
    Dim result As String
    Dim start As Long, ending As Long
    
    start = InStr(html, "The numbers next to the category names below indicate the eBay category number.")
    start = start + Len("The numbers next to the category names below indicate the eBay category number.")
    ending = InStr(start + 1, html, "The numbers next to the category names below indicate the eBay category number.")
    result = Mid(html, start, ending - start + 1)
    CategoriesOnly = result
End Function

Function EliminateGarbage(html As String) As String
    Dim result As String
    result = html
    result = Replace(result, "&nbsp", "")
    result = Replace(result, "<!--L-->", "")
    ' result = Replace(result, "</font>", "")
    result = Replace(result, "</a>", "")
    result = Replace(result, "</i>", "")
    result = Replace(result, "<p>", "")
    result = Replace(result, "</p>", "")
    result = Replace(result, "</b>", "")
    result = Replace(result, "<br>", "")
    'result = Replace(result, "<ul>", "")
    'result = Replace(result, "</ul>", "")
    result = Replace(result, "</i>", "")
    result = Replace(result, "<i>", "")
    result = Replace(result, Chr(10), "")
    result = Replace(result, "font size =", "font size=")
    EliminateGarbage = result
End Function

Function GetCategories(treeViewCtrl As treeview, URL As String) As String
    Dim destFilename As String
    Dim result As String, finalResult As String, tempStr As String
    Dim i As Integer, start As Integer, ending As Integer
    Dim categoryNumber As Long
    Dim categoryName As String
    Dim categoryLevel As Long, categoryDelta As Variant
    Const maxCategories = 20
    Const deltaIncrement = 1
    Dim parentCategoryName(maxCategories) As String
    Dim parentCategoryNumber(maxCategories) As Long
    Dim mNodeSet As Node
    
    result = GetHTTPFile(URL)
    result = CategoriesOnly(result)
    result = EliminateGarbage(result)
    
    i = 1
    categoryDelta = 2#
    Do
        i = InStr(result, "(#")
        If i <> 0 Then
            ending = InStr(i, result, ")")  'With this we find the end of the category #
            'now to get the category name we must travel backwards
            categoryNumber = Val(Mid(result, i + 2, ending - (i + 2) + 1))
            start = InStrRev(result, ">", i)
            categoryName = LTrim(Trim(Mid(result, start + 1, i - start - 1)))
            'If Asc(Mid(categoryName, 1, 1)) = 10 Then
            '  categoryName = Mid(categoryName, 2)
            'End If
            categoryLevel = Val(Mid(result, InStrRev(result, "font size=", i) + 11, 1))
            j = i
            Do
                j = InStrRev(result, "<ul>", j) - 1
                If j >= 1 Then categoryDelta = categoryDelta + deltaIncrement
            Loop While j >= 1
            j = i
            Do
                j = InStrRev(result, "</ul>", j) - 1
                If j >= 1 Then categoryDelta = categoryDelta - deltaIncrement
            Loop While j >= 1
            parentCategoryNumber(categoryDelta) = categoryNumber
            parentCategoryName(categoryDelta) = categoryName
            tempStr = "CL: " & categoryLevel
            tempStr = tempStr & " C#:" & categoryNumber&
            tempStr = tempStr & " CN:" & categoryName
            'tempStr = tempStr & " FC:" & Asc(Mid(categoryName, 1, 1))
            tempStr = tempStr & " CD:" & categoryDelta
            tempStr = tempStr & " CP#:" & parentCategoryNumber(categoryDelta - deltaIncrement)
            tempStr = tempStr & " CPN:" & parentCategoryName(categoryDelta - deltaIncrement)
            Set mNodeSet = treeViewCtrl.Nodes.Add("K" & parentCategoryNumber(categoryDelta - deltaIncrement), tvwChild)
            mNodeSet.Key = "K" & categoryNumber
            mNodeSet.Text = categoryName
            'Debug.Print tempStr
            result = Mid(result, ending + 1)
            finalResult = finalResult + tempStr + Chr(10)
            'i = ending
        End If
    Loop Until i = 0
    GetCategories = finalResult
End Function
