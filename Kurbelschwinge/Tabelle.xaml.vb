Public Class Tabelle
   '--------------------------------------------------------------------------------------------------------------------
   ' DialogFenster Tabelle
   '--------------------------------------------------------------------------------------------------------------------
   ' Verhalten beim Verlassen eines Unterprogramms:
   ' DialogResult = True  --> Dialogfenster wird mit True verlassen
   ' DialogResult = False --> Dialogfenster wird mit False verlassen
   ' DialogResult = Nothing --> Dialogfenster wird nicht verlassen
   '--------------------------------------------------------------------------------------------------------------------
   'Dim DP As New System.Globalization.NumberFormatInfo With {.NumberDecimalSeparator = "."}

   Private Sub ButtonTabelleOK_Click(sender As Object, e As RoutedEventArgs) Handles ButtonTabelleOK.Click
      Dim Fehlermeldung As String = String.Empty   'Kennzeichen für Fehlerfreiheit
      If Not IsNumeric(TextBoxKurbelDrehwinkel.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Drehwinkel fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelX0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Kurbel Drehpunkt X fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelY0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Kurbel Drehpunkt Y fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KurbelLänge fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeX0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Schwinge Drehpunkt X fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeY0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Schwinge Drehpunkt Y fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe SchwingeLänge fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelL1.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelL1 fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelL2.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelL2 fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelLänge fehlerhaft!"
      If Fehlermeldung = String.Empty Then
         '----------------------------------------------------------------------------------------------
         ' Wenn alle Textfelder numerische Eingaben haben, jetzt noch die Umlauffähigkeit prüfen
         '----------------------------------------------------------------------------------------------
         'Dialogfenster temporär speichern
         'Dim KurbelWinkelTemp As Double = TextBoxKurbelDrehwinkel.Text
         Dim KurbelX0Temp As Double = Replace(TextBoxKurbelX0.Text, ".", ",")
         Dim KurbelY0Temp As Double = Replace(TextBoxKurbelY0.Text, ".", ",")
         Dim KurbelLängeTemp As Double = Replace(TextBoxKurbelLänge.Text, ".", ",")
         Dim SchwingeX0Temp As Double = Replace(TextBoxSchwingeX0.Text, ".", ",")
         Dim SchwingeY0Temp As Double = Replace(TextBoxSchwingeY0.Text, ".", ",")
         Dim SchwingeLängeTemp As Double = Replace(TextBoxSchwingeLänge.Text, ".", ",")
         'Dim KoppelL1Temp As Double = TextBoxKoppelL1.Text
         'Dim KoppelL2Temp As Double = TextBoxKoppelL2.Text
         Dim KoppelLängeTemp As Double = Replace(TextBoxKoppelLänge.Text, ".", ",")
         'Abstand der Drehpunkte Kurbel und Schwinge berechnen
         Dim KurbelP0Temp As Point = New Point(KurbelX0Temp, KurbelY0Temp)
         Dim SchwingeP0Temp As Point = New Point(SchwingeX0Temp, SchwingeY0Temp)
         Dim GestellVektor As Vector = Vector.Subtract(SchwingeP0Temp, KurbelP0Temp)
         Dim BasisLängeTemp As Double = GestellVektor.Length
         'Testen, ob das neu eingegebene Getriebe umlauffähig ist
         Dim Bedingung1 As Boolean = (BasisLängeTemp + KurbelLängeTemp) < (KoppelLängeTemp + SchwingeLängeTemp)
         Dim Bedingung2 As Boolean = (BasisLängeTemp + KoppelLängeTemp) > (KurbelLängeTemp + SchwingeLängeTemp)
         Dim Bedingung3 As Boolean = (BasisLängeTemp + SchwingeLängeTemp) > (KurbelLängeTemp + KoppelLängeTemp)
         If Not (Bedingung1 And Bedingung2 And Bedingung3) Then
            Fehlermeldung &= vbCrLf & "Das Getriebe ist nicht umlauffähig!"
         End If
      End If
      'Fehler auswerten
      If Fehlermeldung = String.Empty Then   'kein Fehler vorhanden
         Me.DialogResult = True              'Am Ende dieses UPs wird das Dialogfenster Tabelle mit True geschlossen.
      Else                                   'Fehler vorhanden
         Dim Erklärung As String = Fehlermeldung & vbCrLf & vbCrLf & "Möchten Sie die Tabelle beenden und alle Eingaben verwerfen?"
         If MessageBox.Show(Erklärung, "Eingabefehler", MessageBoxButton.YesNo, MessageBoxImage.Error) = MessageBoxResult.Yes Then
            Me.DialogResult = False          'Am Ende dieses UPs wird das Dialogfenster Tabelle mit False geschlossen.
         End If
         'Hinweis: Wenn Nein gedrückt wurde, dann bleibt das Dialogfenster Tabelle mit DialogResult = Nothing geöffnet.
      End If
   End Sub

   Private Sub ButtonTabelleRunden_Click(sender As Object, e As RoutedEventArgs) Handles ButtonTabelleRunden.Click
      Dim Fehlermeldung As String = String.Empty   'Kennzeichen für Fehlerfreiheit
      If Not IsNumeric(TextBoxKurbelDrehwinkel.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Drehwinkel fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelX0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Kurbel Drehpunkt X fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelY0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Kurbel Drehpunkt Y fehlerhaft!"
      If Not IsNumeric(TextBoxKurbelLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KurbelLänge fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeX0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Schwinge Drehpunkt X fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeY0.Text) Then Fehlermeldung &= vbCrLf & "Eingabe Schwinge Drehpunkt Y fehlerhaft!"
      If Not IsNumeric(TextBoxSchwingeLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe SchwingeLänge fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelL1.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelL1 fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelL2.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelL2 fehlerhaft!"
      If Not IsNumeric(TextBoxKoppelLänge.Text) Then Fehlermeldung &= vbCrLf & "Eingabe KoppelLänge fehlerhaft!"
      If Fehlermeldung = String.Empty Then
         Dim KurbelWinkelTemp As Double = Replace(TextBoxKurbelDrehwinkel.Text, ".", ",")
         Dim KurbelX0Temp As Double = Replace(TextBoxKurbelX0.Text, ".", ",")
         Dim KurbelY0Temp As Double = Replace(TextBoxKurbelY0.Text, ".", ",")
         Dim KurbelLängeTemp As Double = Replace(TextBoxKurbelLänge.Text, ".", ",")
         Dim SchwingeX0Temp As Double = Replace(TextBoxSchwingeX0.Text, ".", ",")
         Dim SchwingeY0Temp As Double = Replace(TextBoxSchwingeY0.Text, ".", ",")
         Dim SchwingeLängeTemp As Double = Replace(TextBoxSchwingeLänge.Text, ".", ",")
         Dim KoppelL1Temp As Double = Replace(TextBoxKoppelL1.Text, ".", ",")
         Dim KoppelL2Temp As Double = Replace(TextBoxKoppelL2.Text, ".", ",")
         Dim KoppelLängeTemp As Double = Replace(TextBoxKoppelLänge.Text, ".", ",")
         TextBoxKurbelDrehwinkel.Text = Math.Round(KurbelWinkelTemp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKurbelX0.Text = Math.Round(KurbelX0Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKurbelY0.Text = Math.Round(KurbelY0Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKurbelLänge.Text = Math.Round(KurbelLängeTemp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxSchwingeX0.Text = Math.Round(SchwingeX0Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxSchwingeY0.Text = Math.Round(SchwingeY0Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxSchwingeLänge.Text = Math.Round(SchwingeLängeTemp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKoppelL1.Text = Math.Round(KoppelL1Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKoppelL2.Text = Math.Round(KoppelL2Temp, Hauptfenster.AnzNachkommastellen).ToString
         TextBoxKoppelLänge.Text = Math.Round(KoppelLängeTemp, Hauptfenster.AnzNachkommastellen).ToString
      Else
         MessageBox.Show(Fehlermeldung, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error)
      End If
      DialogResult = Nothing 'Dialogfenster bleibt geöffnet
   End Sub

   Private Sub ButtonTabelleAbbrechen_Click(sender As Object, e As RoutedEventArgs) 'Handles ButtonTabelleAbbrechen.Click
      Me.DialogResult = False  'schließt das Dialogfenster
   End Sub

   Private Sub CheckBoxKurbelDrehrichtung_Click(sender As Object, e As RoutedEventArgs) 'Handles CheckBoxKurbelDrehrichtung.Click
      If CheckBoxKurbelDrehrichtung.IsChecked Then
         CheckBoxKurbelDrehrichtung.Content = "positiv"
      Else
         CheckBoxKurbelDrehrichtung.Content = "negativ"
      End If
   End Sub
End Class
