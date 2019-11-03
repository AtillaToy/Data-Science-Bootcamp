Attribute VB_Name = "Module11"
Sub TickerTracker()

For Each ws In Worksheets

  ' variable for ticker name
  Dim Ticker As String

  ' ticker count, delta, and volume
  Dim Tcount As Long
  Dim Tdelta As Double
  Dim Volume As Double
  Dim prtchg As Double
  

  Tcount = 0
  Tdelta = 0
  
  'summary table row
  Dim Sumrow As Integer
  Sumrow = 2
  
  

  'Headers
  
  ws.Range("I" & 1).Value = "Ticker"
  ws.Range("J" & 1).Value = "Yearly Change"
  ws.Range("K" & 1).Value = "Percent Change"
  ws.Range("L" & 1).Value = "Total Volume"
    
  'Last row
  Dim Lastrow As Long
  Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

  For i = 2 To Lastrow

    ' Check if same ticker, if different then set values, run calculations and print summary line
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

      ' Set the Ticker Name
      Ticker = ws.Cells(i, 1).Value
      
      ' Finalize tally
      Tcount = Tcount + 1

      ' Finalize volume
      Volume = Volume + ws.Cells(i, 7).Value
      
      
      ' Calcuate Tdelta
      Tdelta = ws.Cells(i, 3).Value - ws.Cells(i - Tcount + 1, 6).Value
      
      ' Calcuate prtchg
      If ws.Cells(i - Tcount + 1, 6).Value <> 0 Then
        prtchg = Tdelta / ws.Cells(i - Tcount + 1, 6).Value
        
        Else
        prtchg = 0
        
        End If
      
      ' Print Ticker name to summary table
      ws.Range("I" & Sumrow).Value = Ticker

      ' Print Volume Amount summary table
      ws.Range("L" & Sumrow).Value = Volume

      ' Print Change with color
      ws.Range("J" & Sumrow).Value = Tdelta
      
        If Tdelta > 0 Then
      
        ws.Range("J" & Sumrow).Interior.ColorIndex = 4
      
        ElseIf Tdelta < 0 Then
        ws.Range("J" & Sumrow).Interior.ColorIndex = 3
      
        End If
      
      ' Print Change Percent, Format Cell
      ws.Range("K" & Sumrow).Value = prtchg
      ws.Range("K" & Sumrow).NumberFormat = "0.00%"
      

      
      ' Next line in summary table
      Sumrow = Sumrow + 1
      
      ' Reset summary values
      Volume = 0
      Tcount = 0
      Tdelta = 0

    Else

      ' Add to the volume
      Volume = Volume + ws.Cells(i, 7).Value
      
      ' Increase Tcount
      Tcount = Tcount + 1
        

    End If

  Next i
  
  
  'find Greatest
  ws.Range("N2").Value = "Biggest Gainer %"
  ws.Range("N3").Value = "Biggest Loser %"
  ws.Range("N4").Value = "Biggest Volume"
  
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"
  
  
  Dim Gpos As Double
  Dim Gposname As String
  Dim Gneg As Double
  Dim Gnegname As String
  Dim GVol As Double
  Dim GVolname As String
  
  Gpos = 0
  Gneg = 1000000
  GVol = 0
  
  Dim Lastrowsum As Integer
  
  'Find lastrow in summary table
  
  Lastrowsum = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
  
  For j = 2 To Lastrowsum
  
    If ws.Cells(j, 11).Value > Gpos Then
        Gpos = ws.Cells(j, 11).Value
        Gposname = ws.Cells(j, 9).Value
        ws.Range("O2").Value = Gposname
        ws.Range("P2").Value = Gpos
        ws.Range("P2").NumberFormat = "0.00%"
    End If
    
    
    If ws.Cells(j, 11).Value < Gneg Then
        Gneg = ws.Cells(j, 11).Value
        Gnegname = ws.Cells(j, 9).Value
        ws.Range("O3").Value = Gnegname
        ws.Range("P3").Value = Gneg
        ws.Range("P3").NumberFormat = "0.00%"
    End If
    
    If ws.Cells(j, 12).Value > GVol Then
        GVol = ws.Cells(j, 12).Value
        GVolname = ws.Cells(j, 9).Value
        ws.Range("O4").Value = GVolname
        ws.Range("P4").Value = GVol
    End If
  
  Next j

Lastrow = 0

Next ws

End Sub

