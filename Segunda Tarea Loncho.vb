Sub Stock()

Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate
    Debug.Print ws.Name
    



    [i1] = "Ticker"
    [J1] = "Yearly Change"
    [k1] = "Percent Change"
    [l1] = "Total Stock Volume"
    [p1] = "Ticker"
    [p1] = "Ticker"
    [q1] = "Value"
    [o2] = "Greatest & Increase"
    [o3] = "Greatest & Decrease"
    [o4] = "Greatest Total Volume"

    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    tIndex = 2
    openvalue = Cells(2, "C").Value
    volumen = Cells(2, "G").Value
    vIndex = 2
    iIndex = 2
    dIndex = 2
    vmindex = 2
    Max = 0
    Min = 0
    Vol = 0
    
  
    
    For i = 2 To lastRow
        If Cells(i, "A") <> Cells(i + 1, "A") Then
            Cells(tIndex, "I") = Cells(i, "A")
            Cells(tIndex, "J") = Cells(i, "F") - openvalue
            Cells(tIndex, "K") = (Cells(i, "F") - openvalue) / openvalue
            tIndex = tIndex + 1
            openvalue = Cells(i + 1, "C")
            volumen = Cells(i + 1, "G")
        Else
        volumen = volumen + Cells(i + 1, "G")
        Cells(tIndex, "L") = volumen
        vIndex = vIndex + 1
        End If
    Next i
    
      
    
 For i = 2 To lastRow
          If Cells(i, "K") > Max Then
            Max = Cells(i, "K")
            Cells(2, "P") = Cells(i, "I")
            Cells(2, "Q") = Max
        End If
    Next i
    
    For i = 2 To lastRow
    
        If Cells(i, "K") < Min Then
            Min = Cells(i, "K")
            Cells(3, "P") = Cells(i, "I")
            Cells(3, "Q") = Min
        End If
        
    Next i
    For i = 2 To lastRow
    
        If Cells(i, "L") > Vol Then
            Vol = Cells(i, "L")
            Cells(4, "P") = Cells(i, "I")
            Cells(4, "Q") = Vol
        End If
    Next i
  
    
   For i = 2 To lastRow
        If Cells(i, "J") < 0 Then
            Cells(i, "J").Interior.ColorIndex = 3
        Else
            Cells(i, "J").Interior.ColorIndex = 4
        End If
    Next i
    
    
    For i = 2 To lastRow
        Cells(i, "K") = FormatPercent(Cells(i, "K"))
    Next i
    
    
        Cells(2, "Q") = FormatPercent(Cells(2, "Q"))
        
        Cells(3, "Q") = FormatPercent(Cells(3, "Q"))
 
    
    
Next ws


   
End Sub



