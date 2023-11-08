# module-2-challenge

I have used the below pseudocode/scaffold to write my code uploaded on Slack by our instructor.
My subroutine is located in the Module.bas file. I was able to save and upload the file as a .bas file and not as .vbs.

Sub stock_analysis():

    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet

    ' Loop through each worksheet (tab) in the Excel file 
    For Each ws In Worksheets
        ' Initialize values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0

        ' Set title row
        ws.Range("I1").Value = ""
        ws.Range("J1").Value = ""
        ws.Range("K1").Value = ""
        ws.Range("L1").Value = ""
        ws.Range("P1").Value = ""
        ws.Range("Q1").Value = ""
        ws.Range("O2").Value = ""
        ws.Range("O3").Value = ""
        ws.Range("O4").Value = ""

        ' get the row number of the last row with data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores results in variables
                ''' Code here 

                ' Handle zero total volume
                If total = 0 Then
                    ''' Code here 

                Else
                    ' Find First non zero starting value
                    ''' Code here 

                    ' Calculate Change
                    change = ''
                    percentChange = ''

                    ' start of the next stock ticker
                    start = i + 1

                    ' print the results
                    ws.Range("I" & 2 + j).Value = ''
                    ws.Range("J" & 2 + j).Value = ''
                    ws.Range("J" & 2 + j).NumberFormat = ''
                    ws.Range("K" & 2 + j).Value = ''
                    ws.Range("K" & 2 + j).NumberFormat = ''
                    ws.Range("L" & 2 + j).Value = ''

                    ' colors positives green and negatives red
                    Select Case change
                        ''' Code here

                    End Select

                End If

                ' reset variables for new stock ticker
                total = ''
                change = ''
                j = ''
                days = ''
                dailyChange = ''

            ' If ticker is still the same add results
            Else
                total = ''

            End If

        Next i

        ' take the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = ''
        ws.Range("Q3") = ''
        ws.Range("Q4") = ''

        ' returns one less because header row not a factor
        increase_number = ''
        decrease_number = ''
        volume_number = ''

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ''
        ws.Range("P3") = ''
        ws.Range("P4") = ''

    Next ws

End Sub
