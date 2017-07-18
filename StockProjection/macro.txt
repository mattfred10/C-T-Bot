

Sub CreateCharts()

    For Each Current In Worksheets
        Worksheets(Current.Name).Activate
        ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
        ActiveChart.SetSourceData Source:=ActiveSheet.Range("$B:$H")
        
        'There is some kind of timing issue with this HasTitle command in excel 2013+
        'Switching it on and off seems to resovle the issue.
        ActiveChart.HasTitle = True
        ActiveChart.HasTitle = False
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.Text = Current.Name
    Next
        
       
End Sub
