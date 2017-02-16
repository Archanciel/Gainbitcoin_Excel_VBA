Attribute VB_Name = "Gainbitcoin"
Option Explicit

Public Function getRateUSDCHF() As Double
'    Application.Volatile True 'enforce recalculation if F9 or at sheet opening
    Dim appIE As Object
    Dim allRowOfData As Object
    Dim rateBidStr As String
    Dim rateAskStr As String
    Dim rateBidDbl As Double
    Dim rateAskDbl As Double
    
    Set appIE = CreateObject("internetexplorer.application")
    
    With appIE
        .Navigate "https://uk.investing.com/currencies/streaming-forex-rates-majors"
        .Visible = False
    End With
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set allRowOfData = appIE.Document.getElementById("pair_4")
    
    rateBidStr = allRowOfData.Cells(2).innerHTML
    rateAskStr = allRowOfData.Cells(3).innerHTML
    
    rateBidDbl = CDbl(rateBidStr)
    rateAskDbl = CDbl(rateAskStr)
    
    appIE.Quit
    Set appIE = Nothing
    
    getRateUSDCHF = (rateBidDbl + rateAskDbl) / 2
End Function

Public Function getRateBTCUSD() As Double
'    Application.Volatile True 'enforce recalculation if F9 or at sheet opening
    Dim appIE As Object
    Dim element As HTMLTableCell
    Dim rateStr As String
    Dim rateDbl As Double
    
    Set appIE = CreateObject("internetexplorer.application")
    
    With appIE
        .Navigate "https://uk.investing.com/currencies/btc-usd"
        .Visible = False
    End With
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set element = appIE.Document.getElementById("lst_49798")
    
    rateStr = Replace(element.innerText, ",", "")
    rateDbl = CDbl(rateStr)
    
    appIE.Quit
    Set appIE = Nothing
    
    getRateBTCUSD = rateDbl
End Function

'Fonction utilisée dans la feuille SIMULATION. Retourne
'le cours du BTC en monnaie currencyStr
'
'Retourne une valeur négative si la monnaie passée en
'paramètre est inconnue
Public Function getRateBTCIn(currencyStr) As Double
    Dim appIE As Object
    Dim dollarCurrency As Double
    Dim usdCurrency As Double
    Dim btcUsd As Double
    
    Set appIE = CreateObject("internetexplorer.application")
    appIE.Visible = False
    
    Select Case currencyStr
        Case "CHF"
            usdCurrency = getRateUSDCHF2(appIE)
        Case "EUR"
            usdCurrency = 0 'getRateUSDEUR()
        Case Else
            usdCurrency = -1
    End Select

    If usdCurrency < 0 Then
        getRateBTCIn = -1
    Else
        btcUsd = getRateBTCUSD2(appIE)
        getRateBTCIn = btcUsd * usdCurrency
    End If
    
    appIE.Quit
    Set appIE = Nothing
End Function

Public Function getRateUSDCHF2(appIE As Object) As Double
'    Application.Volatile True 'enforce recalculation if F9 or at sheet opening
    Dim allRowOfData As Object
    Dim rateBidStr As String
    Dim rateAskStr As String
    Dim rateBidDbl As Double
    Dim rateAskDbl As Double
    
    appIE.Navigate "https://uk.investing.com/currencies/streaming-forex-rates-majors"
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set allRowOfData = appIE.Document.getElementById("pair_4")
    
    rateBidStr = allRowOfData.Cells(2).innerHTML
    rateAskStr = allRowOfData.Cells(3).innerHTML
    
    rateBidDbl = CDbl(rateBidStr)
    rateAskDbl = CDbl(rateAskStr)
    
    getRateUSDCHF2 = (rateBidDbl + rateAskDbl) / 2
End Function

Public Function getRateBTCUSD2(appIE As Object) As Double
'    Application.Volatile True 'enforce recalculation if F9 or at sheet opening
    Dim element As HTMLTableCell
    Dim rateStr As String
    Dim rateDbl As Double
    
    Set appIE = CreateObject("internetexplorer.application")
    
    appIE.Navigate "https://uk.investing.com/currencies/btc-usd"
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set element = appIE.Document.getElementById("lst_49798")
    
    rateStr = Replace(element.innerText, ",", "")
    rateDbl = CDbl(rateStr)
    
    getRateBTCUSD2 = rateDbl
End Function

'Forcing the realtime quotes functions in the range to refetch their value
Sub updateRealTimeRates()
Attribute updateRealTimeRates.VB_ProcData.VB_Invoke_Func = "u\n14"
    Range("COURS_TMPS_REEL").Replace What:="=", Replacement:="=", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub tst()
    Dim tmp As Double
    tmp = getRateBTCIn("CHF")
    MsgBox tmp
End Sub
Private Sub Auto_Open()
    updateRealTimeRates
    MsgBox "Real time rates update successfull (CTRL + U to refresh)", vbOKOnly + vbInformation
End Sub


