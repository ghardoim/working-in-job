Sub replace_info(key As String, value As String, doc_content As Find)
    With doc_content
        .Text = key
        .Replacement.Text = value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
End Sub