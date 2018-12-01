Sub stock_data()

    'Create the Ticker and Total Stock Volume headers

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
    'Hard mode
    
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest $ Decrease"
    Range("N4").Value = "Greatest Total Volume"

    'Create the Yearly Change and Percent Change Headers

    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"

    'Create the Stock_Name and Total_Stock_Volume variables

    Dim Stock_Name As String
    Dim Total_Stock_Volume As Double

    Total_Stock_Volume = 0

    'Create Stock_Row Variable

    Dim Stock_Row As Long

    Stock_Row = 2

    'Create Open_Value and Close_Value Variables

    Dim Open_Value As Double

    Dim Close_Value As Double

    Open_Value = Cells(2, 3).Value

    'Create Yearly_Change and Percent_Change variables

    Dim Yearly_Change As Double

    Dim Percent_Change As Double
    
    'Create Hard Mode Variables
    
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    Dim GPI_Stock as String
    Dim GPD_Stock as String
    Dim GTV_Stock as String

    Greatest_Perecent_Increase = Cells(2, 12).Value
    Greatest_Perecent_Decrease = Cells(2, 12).Value
    Greatest_Total_Volume = Cells(2, 10).Value

    GPI_Stock = Cells(2, 9).Value
    GPD_Stock = Cells(2, 9).Value
    GTV_Stock = Cells(2, 9).Value

    'Find the number of rows

    Dim LastRow As Long

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row


    'Loop through rows

    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            Stock_Name = Cells(i, 1).Value

            Close_Value = Cells(i, 6).Value

            Yearly_Change = Close_Value - Open_Value

            If Open_Value <> 0 Then

                Percent_Change = (Close_Value - Open_Value) / Open_Value

            End If

            Cells(Stock_Row, 9).Value = Stock_Name

            Cells(Stock_Row, 10).Value = Total_Stock_Volume

            Cells(Stock_Row, 11).Value = Yearly_Change

            Cells(Stock_Row, 12).Value = Percent_Change

            Stock_Row = Stock_Row + 1

            Total_Stock_Volume = 0

            Open_Value = Cells(i + 1, 3).Value

        Else

            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If

    Next i

    'Percent Increase Loop

    For i = 3 to LastRow

        If Cells(i, 12).Value > Greatest_Percent_Increase Then

            Greatest_Percent_Increase = Cells(i, 12).Value

            GPI_Stock = Cells(i, 9).Value

        End if

    Next i

    'Percent Decrease Loop

    For i = 3 to LastRow

        If Cells(i, 12).Value < Greatest_Percent_Decrease Then

            Greatest_Percent_Decrease = Cells(i, 12).Value

            GPD_Stock = Cells(i, 9).Value

        End if

    Next i

    'Greatest Total Value Loop

    For i = 3 to LastRow

        If Cells(i, 10).Value > Greatest_Total_Volume Then

            Greatest_Total_Volume = Cells(i, 10).Value

            GTV_Stock = Cells(i, 9).Value

        End if

    Next i

    'Paste Hard Mode

    Range("O2").Value = GPI_Stock
    Range("O3").Value = GPD_Stock
    Range("O4").Value = GTV_Stock

    Range("P2").Value = Greatest_Percent_Increase
    Range("P3").Value = Greatest_Percent_Decrease
    Range("P4").Value = Greatest_Total_Volume


    'Conditional color formatting

    For i = 2 To LastRow

        If Cells(i, 11).Value < 0 Then

            Cells(i, 11).Interior.ColorIndex = 3

        ElseIf Cells(i, 11).Value > 0 Then

            Cells(i, 11).Interior.ColorIndex = 4

        End If

    Next i
    
    'Formatting

    Range("L2:L" & LastRow).NumberFormat = "0.00%"
    Range("K2:K" & LastRow).NumberFormat = "0.00000000"

    Range("P2:P3").NumberFormat = "0.00%"

    Columns("J:P").AutoFit

End Sub


