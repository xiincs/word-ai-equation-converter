Sub Dollar_To_Equation_Both()

    Dim rng As Range
    Set rng = ActiveDocument.Content

    ' ========= 1. 先处理 $$...$$ =========
    With rng.Find
        .ClearFormatting
        .Text = "\$\$*\$\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute

        Dim r1 As Range
        Set r1 = rng.Duplicate

        ActiveDocument.OMaths.Add r1
        ActiveDocument.OMaths(ActiveDocument.OMaths.Count).BuildUp

        rng.Start = r1.End
        rng.End = ActiveDocument.Content.End

    Loop

    ' ========= 2. 再处理 $...$ =========
    Set rng = ActiveDocument.Content

    With rng.Find
        .ClearFormatting
        .Text = "\$*\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute

        Dim r2 As Range
        Set r2 = rng.Duplicate

        ActiveDocument.OMaths.Add r2
        ActiveDocument.OMaths(ActiveDocument.OMaths.Count).BuildUp

        rng.Start = r2.End
        rng.End = ActiveDocument.Content.End

    Loop
    
    ' ========= 3. 统一转专业模式 =========
    Dim eq As OMath
    For Each eq In ActiveDocument.OMaths
        eq.BuildUp
    Next eq

End Sub

