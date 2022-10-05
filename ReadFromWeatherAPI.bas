Attribute VB_Name = "ReadFromWeatherAPI"
Sub Button1_Click()


    Dim XDoc As Object, root As Object
     
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    'API Address
    XDoc.Load ("http://meteo.arso.gov.si/uploads/probase/www/fproduct/text/sl/fcast_SLOVENIA_latest.xml")
    Set root = XDoc.DocumentElement
    
    'Read for which day, 1 is tomorrow, 2 is after tomorrow
    Dim dan As Integer
    dan = 1
    'Delete arrows
    On Error Resume Next
    ActiveSheet.Shapes("Min").Delete
    ActiveSheet.Shapes("Maks").Delete
    'On Error GoTo 0

    On Error GoTo Napaka
    'Get Document Elements
    Set lists = XDoc.DocumentElement
    Set Datum = XDoc.SelectNodes("//metData[" & dan & "]/valid[0]")
    Range("A1") = "Weather data for day " & Datum(0).Text
    Set Vreme = XDoc.SelectNodes("//metData[" & dan & "]/nn_shortText[0]")
    Range("B2") = Vreme(0).Text
    Set MaksTemp = XDoc.SelectNodes("//metData[" & dan & "]/tx[0]")
    Range("B3") = MaksTemp(0).Text & XDoc.SelectNodes("//metData[" & dan & "]/tx_var_unit[0]")(0).Text
    
    'Arrows maximum
    If (CInt(MaksTemp(0).Text) > CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tx[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeUpArrow, 399, 36.75, 10.5, 10.5).Name = "Maks"
    ElseIf (CInt(MaksTemp(0).Text) = CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tx[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, 399, 36.75, 10.5, 10.5).Name = "Maks"
    ElseIf (CInt(MaksTemp(0).Text) < CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tx[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeDownArrow, 399, 36.75, 10.5, 10.5).Name = "Maks"
    End If
    
    Set MinTemp = XDoc.SelectNodes("//metData[" & dan & "]/tn[0]")
    Range("B4") = MinTemp(0).Text & XDoc.SelectNodes("//metData[" & dan & "]/tn_var_unit[0]")(0).Text
    Dim Min As Boolean
    
    'Arrows minimum
    If (CInt(MinTemp(0).Text) > CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tn[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeDownArrow, 399, 51.75, 10.5, 10.5).Name = "Min"
    ElseIf (CInt(MinTemp(0).Text) = CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tn[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, 399, 51.75, 10.5, 10.5).Name = "Min"
    ElseIf (CInt(MinTemp(0).Text) < CInt(XDoc.SelectNodes("//metData[" & dan - 1 & "]/tn[0]")(0).Text)) Then
    ActiveSheet.Shapes.AddShape(msoShapeDownArrow, 399, 51.75, 10.5, 10.5).Name = "Min"
    End If
    
    Set Veter = XDoc.SelectNodes("//metData[" & dan & "]/ff_decodeText_kmh[0]")
    Range("B5") = Veter(0).Text & " km/h"
    'Close the object
    
    Set XDoc = Nothing
    Exit Sub
    
Napaka:
Set XDoc = Nothing
MsgBox ("Error occured while downloading data. Error description:" & vbNewLine & Err.Description)
Range("B2", "B5").Value = 0
End Sub

