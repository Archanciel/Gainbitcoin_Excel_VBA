Attribute VB_Name = "Gainbitcoin"
Option Explicit

'Retourne le cours du $ en CHF en créant et libérant un un objet IE
Public Function getSingleRateUSDCHF() As Double
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
    Set allRowOfData = Nothing
    Set appIE = Nothing
    
    getSingleRateUSDCHF = (rateBidDbl + rateAskDbl) / 2
End Function

'Retourne le cours du BTC en $ en créant et libérant un un objet IE
Public Function getSingleRateBTCUSD() As Double
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
    Set element = Nothing
    Set appIE = Nothing
    
    getSingleRateBTCUSD = rateDbl
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
            usdCurrency = getRate(appIE, "pair_4")
        Case "EUR"
            usdCurrency = 0
        Case Else
            usdCurrency = -1
    End Select

    If usdCurrency < 0 Then
        getRateBTCIn = -1
    ElseIf usdCurrency = 0 Then
        getRateBTCIn = getRate(appIE, "pair_22")
    Else
        btcUsd = getRate(appIE, "pair_21")
        getRateBTCIn = btcUsd * usdCurrency
    End If
    
    appIE.Quit
    Set appIE = Nothing
End Function

Private Function getRateUSDCHF(ByRef appIE As Object) As Double
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
    
'    Set allRowOfData = Nothing
    
    getRateUSDCHF = (rateBidDbl + rateAskDbl) / 2
End Function

Private Function getRateBTCUSD(ByRef appIE As Object) As Double
    Dim element As HTMLTableCell
    Dim rateStr As String
    Dim rateDbl As Double
    
    appIE.Navigate "https://uk.investing.com/currencies/btc-usd"
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set element = appIE.Document.getElementById("lst_49798")
    
    rateStr = Replace(element.innerText, ",", "")
    rateDbl = CDbl(rateStr)
    
'    Set element = Nothing
    
    getRateBTCUSD = rateDbl
End Function

Private Function getRate(ByRef appIE As Object, lineTag As String) As Double
    Dim allRowOfData As Object
    Dim rateBidStr As String
    Dim rateAskStr As String
    Dim rateBidDbl As Double
    Dim rateAskDbl As Double
    
    appIE.Navigate "https://uk.investing.com/currencies/streaming-forex-rates-majors"
    
    Do While appIE.Busy
        DoEvents
    Loop
    
    Set allRowOfData = appIE.Document.getElementById(lineTag)
    
    rateBidStr = allRowOfData.Cells(2).innerHTML
    rateBidStr = Replace(rateBidStr, ",", "")
    rateAskStr = allRowOfData.Cells(3).innerHTML
    rateAskStr = Replace(rateAskStr, ",", "")
    
    rateBidDbl = CDbl(rateBidStr)
    rateAskDbl = CDbl(rateAskStr)
    
'    Set allRowOfData = Nothing
    
    getRate = (rateBidDbl + rateAskDbl) / 2
End Function

'Forcing the realtime quotes functions in the range to refetch their value
Sub updateRealTimeRates()
Attribute updateRealTimeRates.VB_ProcData.VB_Invoke_Func = "U\n14"
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




