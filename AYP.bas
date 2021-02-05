Attribute VB_Name = "Module2"
'Loop from Row 1 -1000000
'Had about four variable to keep track of
'Combine Ticker variable together. by having One type and eloiminate the rest of the same type.
'Plus one of the same type in volumn.
'A=1/B=2/C=3/D=4/E=5/F=6/G=7/H=8
'I=9/J=10/K=11/L=12/M=13/N=14/O=15/P=16
'Q=17/R=18/S=19T/=20/U=21/V=22/W=23/X=24/Y=25

Sub AYP()
'Header
Cells(1, 10).Value = "Ticker"
Cells(1, 18).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Volume"
Cells(1, 14).Value = "Start Price"
Cells(1, 15).Value = "Close Price"
Cells(1, 19).Value = "Value"
Cells(2, 17).Value = "Greatest % Change"
Cells(3, 17).Value = "Lowest % Change"
Cells(4, 17).Value = "Greatest Volumn"
'Datas
    Dim nTicker As String
    Dim cTicker As String
    Dim Ope As Double
    Dim Clo As Double
    Dim YCha As Double
    Dim Pcha As Double
    Dim Vol As Double
    Dim Tick As Integer
' why Tick equal one is to make it start the ways with I
    Tick = 1
    Vol = 0
'we stored the first Ope because we can only get the next one, to generate the next one we look athe end
    Ope = Cells(2, 3).Value
    
'we tried to use Ope the first example before the loop
    'Ope = Cells(2, 3).Value 'A, 57.19, I= 2, Tick = 1
'we tried to use the nex the next example
    'Ope = Cells(263, 3).Value 'AA, 31.89, I = 263, Tick = 2
'we do not want I to continue, we want Tick
'But if Tick continue it will give us the wrong the number
'let tried with string, if the nticker is the same as the first ticker then we ignore it
'let consider another variable
    
'Range formating Percemn
    Range("L:L").NumberFormat = "0.00%"
    
    
    For I = 2 To 1000000
 
'True Statement
        If Cells(I, 1).Value = Cells(I + 1, 1).Value Then
            'This Vol will added 0 + plus the first value before it go the nextone
            Vol = Vol + Cells(I, 7).Value
            'if this is the I is row 2 for "A" then it will be Vol = 0 + 0
            'Test 5
            'Ope = 0
        Else
'False Statement
        'Put info for the current type of Stock
        'The Name, The Volume, Yearly Change, Percentage, Colorcoded
        'Tick will only move if the condition false
            Tick = Tick + 1
        'Name
        
        'Since the first iteration, Tick + 1 = 2. Add the column.
        
            Cells(Tick, 10).Value = Cells(I, 1).Value
            'Ticker A
            
        
        'Volume
        'this vol will add the last value it had because
            Vol = Vol + Cells(I, 7).Value
            Cells(Tick, 13).Value = Vol
            Vol = 0

        'Yearly Change
            'Need to Capture the Clo
            'Clo according to Tick, Last day
            Clo = Cells(I, 6).Value
            Cells(Tick, 15).Value = Clo 'first Tick = 2, I = Continue
            'let's loop through Tick
            ' to showi n the table
            
            'Year Change logic/ current Ticker
            YCha = Clo - Ope
            Cells(Tick, 11).Value = YCha
            
            'percent Change/ Current logic
            'to avoid divide by Zero
            
            If Ope = 0 Then
            Pcha = 1
            Cells(Tick, 12).Value = Pcha
            Else
            Pcha = YCha / Ope
            Cells(Tick, 12).Value = Pcha
            End If
            
            ' Ope in the next Ticker
            Cells(Tick, 14).Value = Ope
            Ope = Cells(I + 1, 3).Value
            
            
            'If Cells(2, 1) == cells(3, 1). Ope1 = cells(2, 3) do not change
            'Ope 1 is showned at Tick,
            'If Cells(262, 1) =\= cells(263, 1). Ope change
            'Ope 2 i showned at the next tick
            
        'Color-Coded
            If Cells(Tick, 12).Value > 0 Then
                Cells(Tick, 12).Interior.ColorIndex = 4
                End If
            If Cells(Tick, 12).Value < 0 Then
                Cells(Tick, 12).Interior.ColorIndex = 3
                End If
        End If
    Next I
                
' Percent Max
    Cells(2, 19).NumberFormat = "0.00%"
    Cells(2, 19).Value = WorksheetFunction.Max(Range("L:L"))


'Percent Min
    'Range("T:T").NumberFormat = "0.00%"
    Cells(3, 19).Value = WorksheetFunction.Min(Range("L:L"))
    Cells(3, 19).NumberFormat = "0.00%"
'Greatest Volumn
    Cells(4, 19).Value = WorksheetFunction.Max(Range("M:M"))

'Use a Varaiable
'Dim PMax As Double
    
    'PMax = Cells(2, 19).Value
'Ticker with Max
    'WorksheetFunction.Lookup
'Dim Lookup As String
    'Lookup = WorksheetFunction.VLookup(Cells(2, 19).Value, Range("J1:M2000"), 1, False)
    'Try it with Conditional

'For anotherloop to look up
For Loo = 2 To 5000
   'Greatest Percentage
   If Cells(Loo, 12).Value = Cells(2, 19).Value Then
    Cells(2, 18).Value = Cells(Loo, 10).Value
   End If
   'Lowest Percentage
   If Cells(Loo, 12).Value = Cells(3, 19).Value Then
    Cells(3, 18).Value = Cells(Loo, 10).Value
   End If
   'Greatest Volumn
   If Cells(Loo, 13).Value = Cells(4, 19).Value Then
    Cells(4, 18).Value = Cells(Loo, 10).Value
   End If
Next Loo
End Sub
