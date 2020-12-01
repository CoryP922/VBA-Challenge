VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub main()
    test ("2014")
    test ("2015")
    test ("2016")
End Sub


Sub test(Worksheet As String)
    Dim ticker As String
    Dim newticker As String
    Dim rng As Range
    Dim change As Double
    Dim openval As Double
    Dim closeval As Double
    Dim volume As Double
   
    Dim WS As Worksheet
    Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
    WS.Name = Worksheet + "-out"
    WS.Cells(1, "A").Value = "Ticker"
    WS.Cells(1, "B").Value = "Change"
    WS.Cells(1, "C").Value = "Percent"
    WS.Cells(1, "D").Value = "Volume"
   
    Dim counter As Integer
    counter = 1
   
    Set rng = ThisWorkbook.Sheets(Worksheet).UsedRange
    'Debug.Print rng.Cells.Count
    For Each looprow In rng.Rows
   
    'Debug.Print
    newticker = looprow.Cells(1, "A").Value
    If newticker = "<ticker>" Then
    'put in the column headers for the new worksheet
    Debug.Print
    Else
        If ticker = newticker Then
        'update the close
            closeval = looprow.Cells(1, "F").Value
           
        ' add the volume to volume
            volume = volume + looprow.Cells(1, "G").Value
        Else
            If ticker <> "" Then
            counter = counter + 1
               
                'output
                ' ticker
                ' calculate the change
                change = (closeval - openval) '/ openval 'or what ever percent function you use
                ' calculate the percent
               
                ' volume
                WS.Cells(counter, "A").Value = ticker
                WS.Cells(counter, "B").Value = change
                If change < 0 Then
                    WS.Cells(counter, "B").Interior.Color = RGB(255, 0, 0)
                Else
                    WS.Cells(counter, "B").Interior.Color = RGB(0, 255, 0)
                End If
                If openval > 0 Then
                    WS.Cells(counter, "C").Value = change / openval
                Else
                    WS.Cells(counter, "C").Value = 0
                End If
                WS.Cells(counter, "D").Value = volume
               
            End If
   
            'reset the values
            openval = looprow.Cells(1, "C").Value
            closeval = looprow.Cells(1, "F").Value
            volume = looprow.Cells(1, "G").Value
            ticker = newticker

        End If
   
    End If
    Next
   
    'output from last group
    ' ticker
    ' calculate the change
    change = (closeval - openval) '/ openval 'or what ever percent function you use
    ' calculate the percent
    ' volume
                WS.Cells(counter, "A").Value = ticker
                WS.Cells(counter, "B").Value = change
                If change < 0 Then
                    WS.Cells(counter, "B").Interior.Color = RGB(255, 0, 0)
                Else
                    WS.Cells(counter, "B").Interior.Color = RGB(0, 255, 0)
                End If
                WS.Cells(counter, "C").Value = change / openval
                WS.Cells(counter, "D").Value = volume
   

End Sub
