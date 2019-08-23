Sub Main
    Dim App As femap.model
    Set App = feFemap()
 
    Dim data As femap.DataTable
    Set data = App.feDataTable
    data.Clear
 
    Dim ElementsSet As femap.Set
    Dim ResultsSet As femap.Set
    Dim OutputvectorsSet As femap.Set
               
    Set ElementsSet = App.feSet
    Set ResultsSet = App.feSet
    Set OutputvectorsSet = App.feSet
 
    ElementsSet.Select( FT_ELEM, True, "Select Elements to Rank the Output" )
    ResultsSet.Select( FT_OUT_CASE, True, "Select Results to Consider" )
               
    OutputvectorsSet.add(1000026) '1000026 - Lam Ply1 Major Principle Stress
    OutputvectorsSet.add(1000027) '1000027 - Lam Ply1 Minor Principle Stress
    OutputvectorsSet.add(1000426) '1000426 - Lam Ply3 Major Principle Stress
    OutputvectorsSet.add(1000427) '1000427 - Lam Ply3 Minor Principle Stress
               
    App.feResultsRankingToDataTable(True,8,1,1,2,1,ElementsSet.ID,ResultsSet.ID,OutputvectorsSet.ID)
               
 
End Sub