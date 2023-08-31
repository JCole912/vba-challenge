Sub stocks_loop()
 
    Dim report As Worksheet
    Dim sheet_names(6) As String
    Dim sheet_number As Integer
    Dim Year_Start_Value As Currency

    Dim Greatest_Pct_Increase_Value As Double
    Dim Greatest_Pct_Decrease_Value As Double
    Dim Greatest_Tot_Volume_Value As LongLong
    
    Dim Greatest_Pct_Increase_Ticker As String
    Dim Greatest_Pct_Decrease_Ticker As String
    Dim Greatest_Tot_Volume_Ticker As String
    
    sheet_names(1) = "A"
    sheet_names(2) = "B"
    sheet_names(3) = "C"
    sheet_names(4) = "D"
    sheet_names(5) = "E"
    sheet_names(6) = "F"
    
    ' Set Dimensions
    Dim i As Long
    Dim J As Integer
    Dim Volume As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double ' Added missing variable
    ' Initialize Variables
    
    For sheet_number = 1 To 6
        Sheets(sheet_names(sheet_number)).Select
        J = 0
        Volume = 0
        
        Greatest_Pct_Increase_Value = 0
        Greatest_Pct_Decrease_Value = 0
        Greatest_Tot_Volume_Value = 0
        
        ' Setting Column Titles
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Change in Percentage"
        Range("M1").Value = "Total Stock Volume"
        
        Range("P2").Value = "Greatest Percentage Increase"
        Range("P3").Value = "Greatest Percentage Decrease"
        Range("P4").Value = "Greatest Total Volume"
        
        Range("Q1").Value = "Ticker"
        Range("R1").Value = "Value"
        
        ' Loop over all rows
        For i = 2 To 22771
            ' If Ticker Change
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                Year_Start_Value = Cells(i, 3).Value
                Volume = Cells(i, 7).Value
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Volume = Volume + Cells(i, 7).Value
                ' Adding Tickers
                Range("J" & (2 + J)).Value = Cells(i, 1).Value
                ' Record Volume
                Range("M" & (2 + J)).Value = Volume
                ' Calculate YearlyChange
                YearlyChange = Cells(i, 6).Value - Year_Start_Value ' Assuming the values are in columns F and C
                Range("K" & (2 + J)).Value = YearlyChange
                ' Calculate PercentageChange
                If Year_Start_Value Then ' Avoid division by zero
                    PercentageChange = YearlyChange / Year_Start_Value
                Else
                    PercentageChange = 0
                End If
                
                If PercentageChange > Greatest_Pct_Increase_Value Then
                    Greatest_Pct_Increase_Value = PercentageChange
                    Greatest_Pct_Increase_Ticker = Cells(i, 1).Value
                End If
                
                If PercentageChange < Greatest_Pct_Decrease_Value Then
                    Greatest_Pct_Decrease_Value = PercentageChange
                    Greatest_Pct_Decrease_Ticker = Cells(i, 1).Value
                End If
                
                If Volume > Greatest_Tot_Volume_Value Then
                    Greatest_Tot_Volume_Value = Volume
                    Greatest_Tot_Volume_Ticker = Cells(i, 1).Value
                End If
                
                Range("L" & 2 + J).Value = PercentageChange
                ' Incrementing J
                J = J + 1
            Else
                ' Add Volume to a Variable
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        ' Write summary info.
        Range("Q2").Value = Greatest_Pct_Increase_Ticker
        Range("Q3").Value = Greatest_Pct_Decrease_Ticker
        Range("Q4").Value = Greatest_Tot_Volume_Ticker
        
        Range("R2").Value = Greatest_Pct_Increase_Value
        Range("R3").Value = Greatest_Pct_Decrease_Value
        Range("R4").Value = Greatest_Tot_Volume_Value
        
        ' Format Summary Table
        Columns("M:M").NumberFormat = "0.00"
    Next sheet_number
    
    MsgBox ("All Done")
End Sub

