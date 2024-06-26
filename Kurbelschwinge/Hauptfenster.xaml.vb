Imports System.Math
'Imports System.Xml


Public Class Hauptfenster

#Region "Variable auf WPF_Hauptfenster-Ebene"

    'Kenngrößen des Getriebe
    Dim KurbelWinkel As Double
    Dim KurbelP0 As Point
    Dim KurbelLänge As Double
    Dim KurbelDrehrichtung As Integer '+1 oder -1
    Dim SchwingeP0 As Point
    Dim SchwingeLänge As Double
    Dim KoppelLänge As Double
    Dim KoppelL1 As Double
    Dim KoppelL2 As Double
    Dim Montage As Integer  '+1 oder -1

    'Folgegrößen zum Zeichnen
    Dim KurbelP1 As Point
    Dim SchwingeP1 As Point
    Dim KoppelP0 As Point
    Dim KoppelP1 As Point
    ReadOnly KoppelpunkteCollection As New PointCollection
    Dim PunktRadius As Double
    ReadOnly Streckglieder As New PointCollection  'Fünf Punkte für die Strecklagen
    Dim PunkteZwischenlagen() As Point

    'Farben
    Dim PinselFenster As SolidColorBrush = New SolidColorBrush(SystemColors.WindowColor)
    Dim PinselGetriebe As SolidColorBrush = New SolidColorBrush(SystemColors.WindowTextColor)
    Dim PinselFüllung As SolidColorBrush = New SolidColorBrush(SystemColors.AppWorkspaceColor)
    Dim PinselKoppel12 As SolidColorBrush = New SolidColorBrush(SystemColors.ActiveCaptionColor)
    Dim PinselStrecklagen As SolidColorBrush = New SolidColorBrush(SystemColors.ActiveCaptionColor)
    Dim PinselZwischenlagen As SolidColorBrush = New SolidColorBrush(SystemColors.InactiveCaptionColor)
    Dim PinselUmFlächen As SolidColorBrush = New SolidColorBrush(Color.FromRgb(204, 204, 255))
    Dim PinselUmLinien As SolidColorBrush = New SolidColorBrush(Color.FromRgb(180, 180, 255))
    Dim Pinseldicke As Double = 1
    Dim PinseldickeFaktor As Double = 1.5
    'Zei-Objekte
    ReadOnly KurbelDrehpunkt As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .Name = "Kurbeldrehpunkt", .ToolTip = .Name}
    ReadOnly SchwingeDrehpunkt As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .Name = "SchwingeDrehpunkt", .ToolTip = .Name}
    ReadOnly KoppelPunkt As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .Name = "KoppelPunkt", .ToolTip = .Name}
    ReadOnly Kurbel As Line = New Line With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke * 3, .Name = "Kurbel", .ToolTip = .Name, .StrokeStartLineCap = PenLineJoin.Round, .StrokeEndLineCap = PenLineCap.Round}
    ReadOnly Schwinge As Line = New Line With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke * 3, .Name = "Schwinge", .ToolTip = .Name, .StrokeStartLineCap = PenLineJoin.Round, .StrokeEndLineCap = PenLineCap.Round}
    ReadOnly Koppel As Line = New Line With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke * 3, .Name = "Koppel", .ToolTip = .Name, .StrokeStartLineCap = PenLineJoin.Round, .StrokeEndLineCap = PenLineCap.Round}
    ReadOnly Koppel1 As Line = New Line With {.Stroke = PinselKoppel12, .StrokeThickness = Pinseldicke * 5, .Name = "KoppelL1", .ToolTip = .Name}
    ReadOnly Koppel2 As Line = New Line With {.Stroke = PinselKoppel12, .StrokeThickness = Pinseldicke * 3, .Name = "KoppelL2", .ToolTip = .Name}
    ReadOnly Koppelkurve As Polygon = New Polygon With {.Stroke = PinselGetriebe, .Name = "Koppelkurve", .ToolTip = .Name}
    'Zei-Objekte Strecklagen
    ReadOnly PLinieStrecklagen As New Polyline With {.Stroke = PinselStrecklagen, .StrokeThickness = Pinseldicke, .Name = "Strecklage", .ToolTip = .Name}
    ReadOnly L1undL2Aussen As New Polyline With {.Stroke = PinselStrecklagen, .StrokeThickness = Pinseldicke, .Name = "L1undL2Aussen", .ToolTip = .Name}
    ReadOnly L1undL2Innen As New Polyline With {.Stroke = PinselStrecklagen, .StrokeThickness = Pinseldicke, .Name = "L1undL2Innen", .ToolTip = .Name}
    'Zei-Objekte Umgebung
    Dim UmLinie As New Line With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .ToolTip = .Name}
    Dim UmRechteckPfad As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .ToolTip = .Name}
    Dim UmKreisPfad As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .ToolTip = .Name}
    Dim UmEllipsePfad As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .ToolTip = .Name}
    Dim UmPolyline As New Polyline With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .ToolTip = .Name}
    Dim UmPolygon As New Polygon With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .ToolTip = .Name}
    Dim UmPfadPfad As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselFüllung, .ToolTip = .Name}
    'Optionen
    Public Shared AnzNachkommastellen As Integer = 2    'MnOptRunden
    Dim AnzPunkteKoppelkurve As Integer = 45  'MnOptKoppelkurve
    Dim Zwischenlagen As Integer = 0          'MnAnsZwischenlagen
    Dim AniSchrittWeite As Double = 2.0       'MnAktion_Schrittweite - Abstand der Phasenbilder in Grad
    Dim AniSchrittZeit As Integer = 20        'MnAktion_Schrittzeit  - Zeit zwischen 2 Phasen in Millisekunden (ms)

    'Programmschalter
    Dim DatenSindVerändert As Boolean = False 'Hauptfenster_Closing
    Dim Strecklagen As Boolean = False        'MnAnsStrecklagen
    Dim GetriebeDreht As Boolean = False      'MnAktion_Bewegen

    'Index in Zei.Children, an dem das Getriebe eingetragen wird
    Dim GetriebeIndex As Integer
    ' Gets a NumberFormatInfo
    ReadOnly DP As New System.Globalization.NumberFormatInfo With {.NumberDecimalDigits = AnzNachkommastellen, .NumberDecimalSeparator = "."}
    'Mauscursor an x,y-Position setzen
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Integer, y As Integer) As Boolean

    ' Kurbelschwinge - Datei
    Dim LwPfadDatei As String = "C:\Ablage\Kurbelschwinge.svg"
    Dim StatusDatei As String  'Kurbelschwinge.xml in My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData

    'Mathematik
    ReadOnly GradBogen As Double = Math.PI / 180.0
    ReadOnly BogenGrad As Double = 180.0 / Math.PI

    ' Mausverhalten
    Enum Aktion
        Keine
        Pan
        ZoomFenster
        KurbelDrehpunkt
        Kurbel
        KurbelWinkel
        SchwingeDrehpunkt
        Schwinge
        KoppelPunkt
        Koppel
        MessageBox
    End Enum
    Dim Maus As Aktion
    Dim Klickpunkt As Point                   'Startpunkt für alle MausAktionen
    Dim TransKlickX, TransKlickY As Double    'Werte beim Mausklick merken

    'Ansicht
    Dim Strecklage1GradKurbel, Strecklage2GradKurbel As Double

    'Zoom
    Dim Transformer As MatrixTransform = New MatrixTransform With {.Matrix = New Matrix(1.0, 0.0, 0.0, -1.0, 0.0, 0.0)}   'M11, M12, M21, M22, OffsetX, OffsetY
    Dim BildLi, BildRe, BildUn, BildOb As Double   'Aktuelle Bild-Maße in mm
    'Zoom Fenster (ziehen mit der rechten Maustaste)
    Dim Fensterpunkt As Point
    ReadOnly Auswahlrahmen As Rectangle = New Rectangle With {.Uid = "Auswahlrechteck", .Stroke = Brushes.Red, .StrokeThickness = 1}

    'Animation durch weiterschalten
    Dim AktionTimer As Timers.Timer = Nothing          'Timer läuft in einem eigenen Thread
    Delegate Sub DelegateZumUI()                       'Delegate für den Zugriffs auf das UI vom TimerThread aus.

#End Region

#Region "UPs auf Klassenebene"

    Private Sub StartgetriebeZeichnen()
        'Kenngrößen des Getriebe
        KurbelWinkel = 135
        KurbelDrehrichtung = 1
        KurbelP0 = New Point(200, 200)
        KurbelLänge = 100

        SchwingeP0 = New Point(700, 100)
        SchwingeLänge = 300

        KoppelLänge = 500
        KoppelL1 = 200
        KoppelL2 = 130
        Montage = 1

        'Optionen
        AnzNachkommastellen = 2    'MnOptRunden
        AnzPunkteKoppelkurve = 45  'MnOptKoppelkurve
        AniSchrittWeite = 2.0      'MnAktion_Schrittweite - Abstand der Phasenbilder in Grad
        AniSchrittZeit = 20        'MnAktion_Schrittzeit  - Zeit zwischen 2 Phasen in Millisekunden (ms)

        'Programmschalter
        DatenSindVerändert = False
        Strecklagen = False
        MenuAnsStrecklagen.IsChecked = False
        Zwischenlagen = 0
        MenuAnsZwischenlagen.IsChecked = False

        '-------------------------------------------------------------------------------------------------------------------------------------
        'Umgebung Linie
        UmLinie = New Line With {.Name = "Linie1", .X1 = "200", .Y1 = "200", .X2 = "700", .Y2 = "100", .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmLinie)

        'Umgebung Rechteck
        UmRechteckPfad = New Path With {.Name = "Rechteck1", .Data = New RectangleGeometry(New Rect(200, 100, 500, 100)), .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = PinselUmFlächen, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmRechteckPfad)

        'Umgebung KreisPfad
        UmKreisPfad = New Path With {.Name = "Kreis1", .Data = New EllipseGeometry() With {.Center = New Point(200, 200), .RadiusX = 100, .RadiusY = 100}, .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = Nothing, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmKreisPfad)

        'Umgebung EllipsePfad
        UmEllipsePfad = New Path With {.Name = "Ellipse1", .Data = New EllipseGeometry() With {.Center = New Point(371, 425), .RadiusX = 67, .RadiusY = 39}, .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = PinselUmFlächen, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmEllipsePfad)

        'Umgebung Polyline
        UmPolyline = New Polyline With {.Name = "Polyline01", .Points = PointCollection.Parse("563.421 377.107 563.421 431 768.418 431 768.418 402.093"), .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = Nothing, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmPolyline)

        'Umgebung Polygon
        UmPolygon = New Polygon With {.Name = "Polygon01", .Points = PointCollection.Parse("95,90 774,90 774,473 95,473"), .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = Nothing, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmPolygon)

        'Umgebung Pfad
        UmPfadPfad = New Path With {.Name = "Pfad1", .Data = New PathGeometry With {.Figures = PathFigureCollection.Parse("M 768.418,402.093 A 299.999 299.999 0 0 1 563.421,377.107")}, .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .ToolTip = .Name, .RenderTransform = Transformer}
        Zei.Children.Add(UmPfadPfad)
        '-------------------------------------------------------------------------------------------------------------------------------------
        ' Ende der Umgebungs-Geometrie. Index merken, an dem die Getriebe-Geometrie beginnt
        GetriebeIndex = Zei.Children.Count
        AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
    End Sub

    Private Sub AllesBerechnenUndWennUmlauffähigNeuZeichnen(ByVal KurP0 As Point, ByVal KurLänge As Double, ByVal SchwiP0 As Point, ByVal SchwiLänge As Double, ByVal KoppLänge As Double)

        'Abstand der Drehpunkte Kurbel und Schwinge berechnen
        Dim GestellVektor As Vector = Vector.Subtract(SchwiP0, KurP0)
        Dim GestellLänge As Double = GestellVektor.Length
        'Testen, ob das neu eingegebene Getriebe umlauffähig ist
        Dim Bedingung1 As Boolean = (GestellLänge + KurLänge) < (KoppLänge + SchwiLänge)
        Dim Bedingung2 As Boolean = (GestellLänge + KoppLänge) > (KurLänge + SchwiLänge)
        Dim Bedingung3 As Boolean = (GestellLänge + SchwiLänge) > (KurLänge + KoppLänge)
        If Bedingung1 And Bedingung2 And Bedingung3 Then
            'das neu eingegebene Getriebe ist umlauffähig -> Argumente an globale Variable übergeben
            'KurbelWinkel = KurbelWinkelTemp
            KurbelP0 = KurP0
            KurbelLänge = KurLänge
            SchwingeP0 = SchwiP0
            SchwingeLänge = SchwiLänge
            'KoppelL1 = KoppelL1Temp
            'KoppelL2 = KoppelL2Temp
            KoppelLänge = KoppLänge
            ''---------------------------------------------------------------------
            ''Liniendicke berechnen
            ''---------------------------------------------------------------------
            'Pinseldicke = Max(SchwingeLänge, KoppelLänge) / 400
            '---------------------------------------------------------------------
            'Strecklagen
            '---------------------------------------------------------------------
            If Strecklagen Then
                Dim GestellLängeQuadrat As Double = GestellVektor.LengthSquared
                Dim GestellBogen As Double = Atan2(GestellVektor.Y, GestellVektor.X)
                '----------------------------------------------------------------------------------------------------------------------------
                'Strecklage Aussen
                '----------------------------------------------------------------------------------------------------------------------------
                Dim Streckglied As Double = KurbelLänge + KoppelLänge
                'Cosinussatz 
                Dim CosinusB1 As Double = (Streckglied * Streckglied + GestellLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * Streckglied * GestellLänge)
                Dim DreieckBogen As Double = Math.Acos(CosinusB1)
                Dim Bogen As Double = GestellBogen + DreieckBogen * Montage
                Strecklage1GradKurbel = Bogen * BogenGrad
                Dim StreckgliedEinheitsVektor As Vector = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
                'Schnittpunkt KurbelP1
                Dim KurbelVektor As Vector = Vector.Multiply(StreckgliedEinheitsVektor, KurbelLänge)
                KurbelP1 = Vector.Add(KurbelVektor, KurbelP0)
                'Schnittpunkt SchwingeP1
                Dim KoppelVektor As Vector = Vector.Multiply(StreckgliedEinheitsVektor, KoppelLänge)
                SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
                'KoppelP0
                Dim KoppelL1Vektor As Vector = Vector.Multiply(StreckgliedEinheitsVektor, KoppelL1)
                KoppelP0 = Vector.Add(KoppelL1Vektor, KurbelP1)
                'KoppelP1
                Dim KoppelL2EinheitsVektor As Vector = New Vector(-StreckgliedEinheitsVektor.Y, StreckgliedEinheitsVektor.X) 'StreckgliedEinheitsVektor um 90° gedreht
                Dim KoppelL2Vektor As Vector = Vector.Multiply(KoppelL2EinheitsVektor, KoppelL2)
                KoppelP1 = Point.Add(KoppelL2Vektor, KoppelP0)
                'PointCollection Streckglieder zur Hälfte füllen
                Streckglieder.Clear()
                Streckglieder.Add(KurbelP0)
                Streckglieder.Add(SchwingeP1)
                Streckglieder.Add(SchwingeP0)
                'PointCollection L1undL2Aussen füllen
                Dim L1undL2PunkteAussen As New PointCollection From {KurbelP1, KoppelP0, KoppelP1}
                '----------------------------------------------------------------------------------------------------------------------------
                'Strecklage innen
                '----------------------------------------------------------------------------------------------------------------------------
                Streckglied = KoppelLänge - KurbelLänge
                'Cosinussatz 
                CosinusB1 = (Streckglied * Streckglied + GestellLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * Streckglied * GestellLänge)
                DreieckBogen = Math.Acos(CosinusB1)
                Bogen = GestellBogen + DreieckBogen * Montage
                Strecklage2GradKurbel = Bogen * BogenGrad + 180
                If Strecklage2GradKurbel > 360 Then Strecklage2GradKurbel -= 360
                StreckgliedEinheitsVektor = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
                'Schnittpunkt KurbelP1
                KurbelVektor = Vector.Multiply(-StreckgliedEinheitsVektor, KurbelLänge)
                KurbelP1 = Vector.Add(KurbelVektor, KurbelP0)
                'Schnittpunkt SchwingeP1
                KoppelVektor = Vector.Multiply(StreckgliedEinheitsVektor, KoppelLänge)
                SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
                'KoppelP0
                KoppelL1Vektor = Vector.Multiply(StreckgliedEinheitsVektor, KoppelL1)
                KoppelP0 = Vector.Add(KoppelL1Vektor, KurbelP1)
                'KoppelP1
                KoppelL2EinheitsVektor = New Vector(-StreckgliedEinheitsVektor.Y, StreckgliedEinheitsVektor.X) 'StreckgliedEinheitsVektor um 90° gedreht
                KoppelL2Vektor = Vector.Multiply(KoppelL2EinheitsVektor, KoppelL2)
                KoppelP1 = Point.Add(KoppelL2Vektor, KoppelP0)
                'PointCollection Streckglieder zweite Hälfte füllen
                Streckglieder.Add(SchwingeP1)
                Streckglieder.Add(KurbelP1)
                'PointCollection L1undL2Innen füllen
                Dim L1undL2PunkteInnen As New PointCollection From {KurbelP1, KoppelP0, KoppelP1}
                'Polylinien mit PointCollections füllen
                PLinieStrecklagen.Points = Streckglieder
                PLinieStrecklagen.StrokeThickness = Pinseldicke
                PLinieStrecklagen.RenderTransform = Transformer
                L1undL2Aussen.Points = L1undL2PunkteAussen
                L1undL2Aussen.StrokeThickness = Pinseldicke
                L1undL2Aussen.RenderTransform = Transformer
                L1undL2Innen.Points = L1undL2PunkteInnen
                L1undL2Innen.StrokeThickness = Pinseldicke
                L1undL2Innen.RenderTransform = Transformer
            End If
            '---------------------------------------------------------------------
            'Zwischenlagen
            '---------------------------------------------------------------------
            For i As Integer = 0 To Zwischenlagen - 1
                Dim AktKurbelwinkel As Double = KurbelWinkel + 360.0 * i / Zwischenlagen
                '---------------------------------------------------
                'Berechnen: KoppelP1
                '---------------------------------------------------
                Dim Drehung As New RotateTransform With {.Angle = AktKurbelwinkel}
                KurbelP1 = KurbelP0 + Drehung.Transform(New Vector(KurbelLänge, 0))
                'Folgegrößen Basis
                Dim BasisVektor As Vector = Vector.Subtract(SchwingeP0, KurbelP1)
                Dim BasisLänge As Double = BasisVektor.Length
                Dim BasisLängeQuadrat As Double = BasisVektor.LengthSquared
                Dim BasisBogen As Double = Atan2(BasisVektor.Y, BasisVektor.X)
                'Grenzlagen behandeln, else allgemeinen Fall
                Dim Bogen As Double
                If BasisLänge = (KoppelLänge + SchwingeLänge) Then          'Strecklage
                    Bogen = BasisBogen
                ElseIf KoppelLänge = (BasisLänge + SchwingeLänge) Then      'Strecklage
                    Bogen = BasisBogen
                ElseIf SchwingeLänge = (BasisLänge + KoppelLänge) Then      'Strecklage
                    Bogen = BasisBogen + Math.PI
                Else                                                        'allgemeiner Fall
                    'Cosinussatz 
                    Dim CosinusB1 As Double = (KoppelLänge * KoppelLänge + BasisLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * KoppelLänge * BasisLänge)
                    Dim DreieckBogen As Double = Math.Acos(CosinusB1)
                    Bogen = BasisBogen + DreieckBogen * Montage
                End If
                Dim KoppelEinheitsVektor As Vector = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
                'KoppelP0
                Dim KoppelL1Vektor As Vector = Vector.Multiply(KoppelEinheitsVektor, KoppelL1)
                KoppelP0 = Vector.Add(KoppelL1Vektor, KurbelP1)
                'KoppelP1
                Dim KoppelL2EinheitsVektor As Vector = New Vector(-KoppelEinheitsVektor.Y, KoppelEinheitsVektor.X) 'KoppelEinheitsVektor um 90° gedreht
                Dim KoppelL2Vektor As Vector = Vector.Multiply(KoppelL2EinheitsVektor, KoppelL2)
                KoppelP1 = Point.Add(KoppelL2Vektor, KoppelP0)
                PunkteZwischenlagen(i) = KoppelP1
            Next
            '---------------------------------------------------------------------
            'Koppelpunkte-Collection leeren und mit neu berechneten Punkten füllen
            'Lage und Größe der Getriebeglieder berechnen
            '---------------------------------------------------------------------
            'Größe der Punkte
            PunktRadius = Min(KurbelLänge, SchwingeLänge) * 0.05
            KoppelpunkteCollection.Clear()
            Dim GradAbstand As Integer = 360 / AnzPunkteKoppelkurve
            For i As Integer = 1 To AnzPunkteKoppelkurve
                Dim Drehung As New RotateTransform With {.Angle = KurbelWinkel + i * GradAbstand}
                KurbelP1 = KurbelP0 + Drehung.Transform(New Vector(KurbelLänge, 0))
                'Folgegrößen Basis
                Dim BasisVektor As Vector = Vector.Subtract(SchwingeP0, KurbelP1)
                Dim BasisLänge As Double = BasisVektor.Length
                Dim BasisLängeQuadrat As Double = BasisVektor.LengthSquared
                Dim BasisBogen As Double = Atan2(BasisVektor.Y, BasisVektor.X)
                'Fehlkonstruktionen ausschließen
                If BasisLänge > (KoppelLänge + SchwingeLänge) Then Continue For
                If KoppelLänge > (BasisLänge + SchwingeLänge) Then Continue For
                If SchwingeLänge > (BasisLänge + KoppelLänge) Then Continue For
                'Grenzlagen behandeln, else allgemeinen Fall
                Dim Bogen As Double
                If BasisLänge = (KoppelLänge + SchwingeLänge) Then          'Strecklage
                    Bogen = BasisBogen
                ElseIf KoppelLänge = (BasisLänge + SchwingeLänge) Then      'Strecklage
                    Bogen = BasisBogen
                ElseIf SchwingeLänge = (BasisLänge + KoppelLänge) Then      'Strecklage
                    Bogen = BasisBogen + Math.PI
                Else                                                        'allgemeiner Fall
                    'Cosinussatz 
                    Dim CosinusB1 As Double = (KoppelLänge * KoppelLänge + BasisLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * KoppelLänge * BasisLänge)
                    Dim DreieckBogen As Double = Math.Acos(CosinusB1)
                    Bogen = BasisBogen + DreieckBogen * Montage
                End If
                'Schnittpunkt SchwingeP1
                Dim KoppelEinheitsVektor As Vector = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
                Dim KoppelVektor As Vector = Vector.Multiply(KoppelEinheitsVektor, KoppelLänge)
                SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
                'KoppelP0
                Dim KoppelL1Vektor As Vector = Vector.Multiply(KoppelEinheitsVektor, KoppelL1)
                KoppelP0 = Vector.Add(KoppelL1Vektor, KurbelP1)
                'KoppelP1
                Dim KoppelL2EinheitsVektor As Vector = New Vector(-KoppelEinheitsVektor.Y, KoppelEinheitsVektor.X) 'KoppelEinheitsVektor um 90° gedreht
                Dim KoppelL2Vektor As Vector = Vector.Multiply(KoppelL2EinheitsVektor, KoppelL2)
                KoppelP1 = Point.Add(KoppelL2Vektor, KoppelP0)
                KoppelpunkteCollection.Add(KoppelP1)
            Next
            '---------------------------------------------------------------------
            ' Geometrieelemente füllen
            '---------------------------------------------------------------------
            'Kurbel Drehpunkt
            KurbelDrehpunkt.Data = New EllipseGeometry() With {.Center = KurbelP0, .RadiusX = PunktRadius, .RadiusY = PunktRadius}
            KurbelDrehpunkt.StrokeThickness = Pinseldicke
            KurbelDrehpunkt.RenderTransform = Transformer
            'Kurbel
            Kurbel.X1 = KurbelP0.X
            Kurbel.Y1 = KurbelP0.Y
            Kurbel.X2 = KurbelP1.X
            Kurbel.Y2 = KurbelP1.Y
            Kurbel.StrokeThickness = Pinseldicke * 3
            Kurbel.RenderTransform = Transformer
            'Schwinge Drehpunkt
            SchwingeDrehpunkt.Data = New EllipseGeometry() With {.Center = SchwingeP0, .RadiusX = PunktRadius, .RadiusY = PunktRadius}
            SchwingeDrehpunkt.StrokeThickness = Pinseldicke
            SchwingeDrehpunkt.RenderTransform = Transformer
            'Schwinge
            Schwinge.X1 = SchwingeP0.X
            Schwinge.Y1 = SchwingeP0.Y
            Schwinge.X2 = SchwingeP1.X
            Schwinge.Y2 = SchwingeP1.Y
            Schwinge.StrokeThickness = Pinseldicke * 3
            Schwinge.RenderTransform = Transformer
            'Koppel
            Koppel.X1 = KurbelP1.X
            Koppel.Y1 = KurbelP1.Y
            Koppel.X2 = SchwingeP1.X
            Koppel.Y2 = SchwingeP1.Y
            Koppel.StrokeThickness = Pinseldicke * 3
            Koppel.RenderTransform = Transformer
            'Koppel1
            Koppel1.X1 = KurbelP1.X
            Koppel1.Y1 = KurbelP1.Y
            Koppel1.X2 = KoppelP0.X
            Koppel1.Y2 = KoppelP0.Y
            Koppel1.StrokeThickness = Pinseldicke * 5
            Koppel1.RenderTransform = Transformer
            'Koppel2
            Koppel2.X1 = KoppelP0.X
            Koppel2.Y1 = KoppelP0.Y
            Koppel2.X2 = KoppelP1.X
            Koppel2.Y2 = KoppelP1.Y
            Koppel2.StrokeThickness = Pinseldicke * 3
            Koppel2.RenderTransform = Transformer
            'Koppelpunkt
            KoppelPunkt.Data = New EllipseGeometry() With {.Center = KoppelP1, .RadiusX = PunktRadius, .RadiusY = PunktRadius}
            KoppelPunkt.StrokeThickness = Pinseldicke
            KoppelPunkt.RenderTransform = Transformer
            'Koppelkurve
            Koppelkurve.Points = KoppelpunkteCollection
            Koppelkurve.StrokeThickness = Pinseldicke
            Koppelkurve.RenderTransform = Transformer
            '---------------------------------------------------------------------
            'Alle Zei.Children nach der Umgebung löschen und neu füllen
            '---------------------------------------------------------------------
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)
            'Zei.Children.Clear()

            'Strecklagen
            If Strecklagen Then
                Zei.Children.Add(PLinieStrecklagen)
                Zei.Children.Add(L1undL2Aussen)
                Zei.Children.Add(L1undL2Innen)
            End If
            'Zwischenlagen
            For i = 0 To Zwischenlagen - 1
                Dim ZwischenPunkt As New Path With {.Stroke = PinselGetriebe, .StrokeThickness = Pinseldicke, .Fill = PinselZwischenlagen} 'With {.Stroke = Brushes.Blue, .StrokeThickness = 1, .Fill = SystemColors.ActiveBorderBrush}
                ZwischenPunkt.Data = New EllipseGeometry() With {.Center = PunkteZwischenlagen(i), .RadiusX = PunktRadius, .RadiusY = PunktRadius}
                Dim Name As String = "Zwischenpunkt" & i.ToString
                ZwischenPunkt.Name = Name
                ZwischenPunkt.ToolTip = Name
                ZwischenPunkt.RenderTransform = Transformer
                Zei.Children.Add(ZwischenPunkt)
            Next
            'Kurbelschwinge in Nennlage
            Zei.Children.Add(Koppel1)
            Zei.Children.Add(Koppel2)

            Zei.Children.Add(Kurbel)
            Zei.Children.Add(Schwinge)
            Zei.Children.Add(Koppel)

            Zei.Children.Add(Koppelkurve)
            Zei.Children.Add(KoppelPunkt)
            Zei.Children.Add(KurbelDrehpunkt)
            Zei.Children.Add(SchwingeDrehpunkt)
        Else
            MessageBox.Show("Das neu eingegebene Getriebe ist nicht umlauffähig.")
        End If
    End Sub

    Private Sub XML_SVG_speichern()
        'SVG-Farbformat  fill="#FF0000" fill="rgb(255, 0, 0)" fill="rgb(100%, 0%, 0%)"
        Dim FarbeFenster As String = "rgb(" & PinselFenster.Color.R & ", " & PinselFenster.Color.G & ", " & PinselFenster.Color.B & ")"
        Dim FarbeGetriebe As String = "rgb(" & PinselGetriebe.Color.R & ", " & PinselGetriebe.Color.G & ", " & PinselGetriebe.Color.B & ")"
        Dim FarbeFüllung As String = "rgb(" & PinselFüllung.Color.R & ", " & PinselFüllung.Color.G & ", " & PinselFüllung.Color.B & ")"
        Dim FarbeKoppel12 As String = "rgb(" & PinselKoppel12.Color.R & ", " & PinselKoppel12.Color.G & ", " & PinselKoppel12.Color.B & ")"
        Dim FarbeStrecklagen As String = "rgb(" & PinselStrecklagen.Color.R & ", " & PinselStrecklagen.Color.G & ", " & PinselStrecklagen.Color.B & ")"
        Dim FarbeZwischenlagen As String = "rgb(" & PinselZwischenlagen.Color.R & ", " & PinselZwischenlagen.Color.G & ", " & PinselZwischenlagen.Color.B & ")"
        Dim FarbeUmFlächen As String = "rgb(" & PinselUmFlächen.Color.R & ", " & PinselUmFlächen.Color.G & ", " & PinselUmFlächen.Color.B & ")"
        Dim FarbeUmLinien As String = "rgb(" & PinselUmLinien.Color.R & ", " & PinselUmLinien.Color.G & ", " & PinselUmLinien.Color.B & ")"
        '--------------------------------------------------------------------------------------------
        ' Grenzwerte für ViewBox ermitteln -> BildLi, BildRe, BildUn, BildOb (wie ZoomAlles)
        '--------------------------------------------------------------------------------------------
        BildLi = Double.MaxValue
        BildRe = Double.MinValue
        BildUn = Double.MaxValue
        BildOb = Double.MinValue
        'Dim t As String = ""
        'Schleife durch alle Zei.Children
        For i As Integer = 0 To Zei.Children.Count - 1
            Select Case Zei.Children(i).GetType.ToString
                Case "System.Windows.Shapes.Line"
                    Dim Linie As Line = Zei.Children(i)
                    BildLi = Math.Min(Linie.X1, BildLi)
                    BildRe = Math.Max(Linie.X1, BildRe)
                    BildUn = Math.Min(Linie.Y1, BildUn)
                    BildOb = Math.Max(Linie.Y1, BildOb)
                    BildLi = Math.Min(Linie.X2, BildLi)
                    BildRe = Math.Max(Linie.X2, BildRe)
                    BildUn = Math.Min(Linie.Y2, BildUn)
                    BildOb = Math.Max(Linie.Y2, BildOb)
                Case "System.Windows.Shapes.Polyline"
                    Dim Polylinie As Polyline = Zei.Children(i)
                    For Each Punkt As Point In Polylinie.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Polygon"
                    Dim Vieleck As Polygon = Zei.Children(i)
                    For Each Punkt As Point In Vieleck.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Path"
                    Dim Pfad As Path = Zei.Children(i)
                    Dim Geo As Geometry = Pfad.Data
                    Dim Grenzen As Rect = Geo.Bounds
                    BildLi = Math.Min(Grenzen.Left, BildLi)
                    BildRe = Math.Max(Grenzen.Right, BildRe)
                    BildUn = Math.Min(Grenzen.Top, BildUn)
                    BildOb = Math.Max(Grenzen.Bottom, BildOb)
            End Select
            't &= i & " " & " L=" & BildLi & " R=" & BildRe & " U=" & BildUn & " O=" & BildOb & vbCrLf
        Next
        'MessageBox.Show(t)
        'Noch einen Rand für den Koppelpunkt
        BildLi -= PunktRadius
        BildRe += PunktRadius
        BildUn -= PunktRadius
        BildOb += PunktRadius
        'MessageBox.Show("BildLi, BildRe, BildUn, BildOb: " & BildLi & "  " & BildRe & "  " & BildUn & "  " & BildOb)
        Dim ViewBox As Rect = New Rect With {.X = BildLi, .Y = -BildOb, .Width = BildRe - BildLi, .Height = BildOb - BildUn}
        'MessageBox.Show(ViewBox.ToString(DP))
        '--------------------------------------------------------------------------------------------
        ' XML-SVG-Datei anlegen und Kopf schreiben
        '--------------------------------------------------------------------------------------------
        Dim Settings As New Xml.XmlWriterSettings() With {.Indent = True, .Encoding = System.Text.Encoding.Unicode}
        Using Schreiber As Xml.XmlWriter = Xml.XmlTextWriter.Create(LwPfadDatei, Settings)
            Schreiber.WriteStartDocument()           '<?xml version="1.0" encoding="utf-8"?>
            Schreiber.WriteComment("Daten von und für das Programm Kurbelschwinge von Wilfried Seibt.")
            Schreiber.WriteDocType("svg", "-//W3C//DTD SVG 1.1//EN", "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd", vbNullString)  'Angabe der korrekten DTD (hier SVG Version 1.1)
            Schreiber.WriteStartElement("svg", "http://www.w3.org/2000/svg")
            Schreiber.WriteAttributeString("viewBox", ViewBox.ToString(DP))
            Schreiber.WriteElementString("title", "Kurbelschwinge von Wilfried Seibt")
            'Rechteck mit Füllung in der Hintergrundfarbe
            Schreiber.WriteStartElement("rect")
            Schreiber.WriteAttributeString("Name", "Hintergrundfarbe")
            Schreiber.WriteAttributeString("x", BildLi.ToString(DP))
            Schreiber.WriteAttributeString("y", (-BildOb).ToString(DP))
            Schreiber.WriteAttributeString("width", (BildRe - BildLi).ToString(DP))
            Schreiber.WriteAttributeString("height", (BildOb - BildUn).ToString(DP))
            Schreiber.WriteAttributeString("fill", FarbeFenster)
            Schreiber.WriteEndElement()  '</rect>
            Schreiber.WriteStartElement("desc")
            '--------------------------------------------------------------------------------------------
            ' Kurbelschleife Getriebe
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteStartElement("Getriebe")
            'Kurbel
            Schreiber.WriteStartElement("Kurbel")
            Schreiber.WriteAttributeString("Drehpunkt", KurbelP0.ToString(DP))
            Schreiber.WriteAttributeString("Länge", KurbelLänge.ToString(DP))
            Schreiber.WriteAttributeString("Drehwinkel", KurbelWinkel.ToString(DP))
            Schreiber.WriteAttributeString("Drehrichtung", KurbelDrehrichtung.ToString(DP))
            Schreiber.WriteEndElement()  '</Kurbel>
            'Schwinge
            Schreiber.WriteStartElement("Schwinge")
            Schreiber.WriteAttributeString("Drehpunkt", SchwingeP0.ToString(DP))
            Schreiber.WriteAttributeString("Länge", SchwingeLänge.ToString(DP))
            Schreiber.WriteEndElement()   '</Schwinge>
            'Koppel
            Schreiber.WriteStartElement("Koppel")
            Schreiber.WriteAttributeString("Länge", KoppelLänge.ToString(DP))
            Schreiber.WriteAttributeString("Länge1", KoppelL1.ToString(DP))
            Schreiber.WriteAttributeString("Länge2", KoppelL2.ToString(DP))
            Schreiber.WriteAttributeString("Montage", Montage.ToString)
            Schreiber.WriteEndElement()   '</Koppel>
            '</Getriebe>
            Schreiber.WriteEndElement()
            'Folgegrößen
            Schreiber.WriteStartElement("Folgegrößen")
            Schreiber.WriteStartElement("Kurbel")
            Schreiber.WriteElementString("KurbelP1", KurbelP1.ToString(DP))
            Schreiber.WriteEndElement()   '</Kurbel>
            Schreiber.WriteStartElement("Schwinge")
            Schreiber.WriteElementString("SchwingeP1", SchwingeP1.ToString(DP))
            Schreiber.WriteEndElement()   '</Schwinge>
            Schreiber.WriteStartElement("Koppel")
            Schreiber.WriteElementString("KoppelP0", KoppelP0.ToString(DP))
            Schreiber.WriteElementString("KoppelP1", KoppelP1.ToString(DP))
            Schreiber.WriteEndElement()   '</Koppel>
            Schreiber.WriteEndElement()   '</Folgegrößen>
            'Programm
            Schreiber.WriteStartElement("Programm")
            Schreiber.WriteStartElement("Optionen")
            Schreiber.WriteAttributeString("Strecklagen", Strecklagen.ToString(DP))
            Schreiber.WriteAttributeString("AnzahlZwischenlagen", Zwischenlagen.ToString(DP))
            Schreiber.WriteAttributeString("AnzPunkteKoppelkurve", AnzPunkteKoppelkurve.ToString(DP))
            Schreiber.WriteEndElement()   '</Optionen>
            Schreiber.WriteStartElement("Aktion")
            Schreiber.WriteAttributeString("AniSchrittWeite", AniSchrittWeite.ToString(DP))
            Schreiber.WriteAttributeString("AniSchrittZeit", AniSchrittZeit.ToString(DP))
            Schreiber.WriteEndElement()   '</Aktion>
            Schreiber.WriteEndElement()   '</Programm>
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteEndElement()   '</desc>
            '--------------------------------------------------------------------------------------------
            ' SVG-Bild Geometrie der Umgebung
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteStartElement("g")
            Schreiber.WriteAttributeString("id", "Umgebung")
            Schreiber.WriteAttributeString("stroke", FarbeUmLinien)
            Schreiber.WriteAttributeString("stroke-width", "1")
            Schreiber.WriteAttributeString("transform", "scale(1,-1)")
            'Schleife durch alle Zei-Kinder GetriebeIndex
            For iZeiKind As Integer = 0 To GetriebeIndex - 1
                Dim aktShape As System.Windows.Shapes.Shape = Zei.Children(iZeiKind)
                Select Case aktShape.ToString
                    Case "System.Windows.Shapes.Line"
                        Dim aktLinie As Line = Zei.Children(iZeiKind)
                        Schreiber.WriteStartElement("line")
                        Schreiber.WriteAttributeString("id", aktLinie.Name)
                        Schreiber.WriteAttributeString("x1", aktLinie.X1.ToString(DP))
                        Schreiber.WriteAttributeString("y1", aktLinie.Y1.ToString(DP))
                        Schreiber.WriteAttributeString("x2", aktLinie.X2.ToString(DP))
                        Schreiber.WriteAttributeString("y2", aktLinie.Y2.ToString(DP))
                        Schreiber.WriteEndElement()  '</line>
                    Case "System.Windows.Shapes.Polyline"
                        Dim aktPolylinie As Polyline = Zei.Children(iZeiKind)
                        Schreiber.WriteStartElement("polyline")
                        Schreiber.WriteAttributeString("id", aktPolylinie.Name)
                        Schreiber.WriteAttributeString("points", aktPolylinie.Points.ToString(DP))
                        Schreiber.WriteAttributeString("fill", "none")
                        Schreiber.WriteEndElement()  '</polyline>
                    Case "System.Windows.Shapes.Polygon"
                        Dim aktPolygon As Polygon = Zei.Children(iZeiKind)
                        Schreiber.WriteStartElement("polygon")
                        Schreiber.WriteAttributeString("id", aktPolygon.Name)
                        Schreiber.WriteAttributeString("points", aktPolygon.Points.ToString(DP))
                        Schreiber.WriteAttributeString("fill", If(IsNothing(aktPolygon.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                        Schreiber.WriteEndElement()  '</polygon>
                    Case "System.Windows.Shapes.Path"
                        Dim aktPfad As Path = Zei.Children(iZeiKind)
                        Select Case aktPfad.Data.GetType.ToString
                            Case "System.Windows.Media.RectangleGeometry"
                                Dim aktRechteckPfad As Path = Zei.Children(iZeiKind)
                                Dim RechteckGeo As RectangleGeometry = aktRechteckPfad.Data
                                Dim Rechteck As Rect = RechteckGeo.Rect
                                Schreiber.WriteStartElement("rect")
                                Schreiber.WriteAttributeString("id", aktRechteckPfad.Name)
                                Schreiber.WriteAttributeString("x", Rechteck.X)
                                Schreiber.WriteAttributeString("y", Rechteck.Y)
                                Schreiber.WriteAttributeString("width", Rechteck.Width)
                                Schreiber.WriteAttributeString("height", Rechteck.Height)
                                Schreiber.WriteAttributeString("fill", If(IsNothing(aktRechteckPfad.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                                Schreiber.WriteEndElement()  '</rect>
                            Case "System.Windows.Media.EllipseGeometry"     'Kreise und Ellipsen
                                Dim aktEllipsePfad As Path = Zei.Children(iZeiKind)
                                Dim EllipseGeo As EllipseGeometry = aktEllipsePfad.Data
                                If EllipseGeo.RadiusX = EllipseGeo.RadiusY Then
                                    'Kreis eintragen
                                    Schreiber.WriteStartElement("circle")
                                    Schreiber.WriteAttributeString("id", aktEllipsePfad.Name)
                                    Schreiber.WriteAttributeString("cx", EllipseGeo.Center.X.ToString(DP))
                                    Schreiber.WriteAttributeString("cy", EllipseGeo.Center.Y.ToString(DP))
                                    Schreiber.WriteAttributeString("r", EllipseGeo.RadiusX)
                                    Schreiber.WriteAttributeString("fill", If(IsNothing(aktEllipsePfad.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                                    Schreiber.WriteEndElement()  '</circle>
                                Else
                                    'Ellipse eintragen
                                    Schreiber.WriteStartElement("ellipse")
                                    Schreiber.WriteAttributeString("id", aktEllipsePfad.Name)
                                    Schreiber.WriteAttributeString("cx", EllipseGeo.Center.X.ToString(DP))
                                    Schreiber.WriteAttributeString("cy", EllipseGeo.Center.Y.ToString(DP))
                                    Schreiber.WriteAttributeString("rx", EllipseGeo.RadiusX)
                                    Schreiber.WriteAttributeString("ry", EllipseGeo.RadiusY)
                                    Schreiber.WriteAttributeString("fill", If(IsNothing(aktEllipsePfad.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                                    Schreiber.WriteEndElement()  '</ellipse>
                                End If
                            Case "System.Windows.Media.PathGeometry"
                                Dim aktPfadPfad As Path = Zei.Children(iZeiKind)
                                Dim PfadGeo As PathGeometry = aktPfadPfad.Data
                                Dim DataString As String = PfadGeo.ToString(DP)
                                Schreiber.WriteStartElement("path")
                                Schreiber.WriteAttributeString("id", aktPfadPfad.Name)
                                Schreiber.WriteAttributeString("d", DataString)
                                Schreiber.WriteAttributeString("fill", If(IsNothing(aktPfadPfad.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                                Schreiber.WriteEndElement()  '</path>
                        End Select
                End Select
            Next ' iZeiKind
            '--------------------------------------------------------------------------------------------
            ' SVG-Bild Geometrie der Umgebung geschrieben
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteEndElement()   '</g>
            '--------------------------------------------------------------------------------------------
            ' SVG-Bild Geometrie des Getriebes
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteStartElement("g")
            Schreiber.WriteAttributeString("transform", "scale(1,-1)")
            'Strecklagen zeichnen, falls angezeigt
            If Strecklagen Then
                Dim Punktestring As String = Streckglieder(0).ToString(DP) & " " & Streckglieder(1).ToString(DP) & " " & Streckglieder(2).ToString(DP) & " " & Streckglieder(3).ToString(DP) & " " & Streckglieder(4).ToString(DP)
                Schreiber.WriteStartElement("polyline")
                Schreiber.WriteAttributeString("points", Punktestring)
                Schreiber.WriteAttributeString("stroke", FarbeStrecklagen)
                Schreiber.WriteAttributeString("stroke-width", "1")
                Schreiber.WriteAttributeString("fill", "none")
                Schreiber.WriteAttributeString("id", "Strecklagen")
                Schreiber.WriteEndElement()  '</Strecklagen>
            End If
            'Zwischenlagen zeichnen, falls angezeigt
            If Zwischenlagen Then
                For i As Integer = 0 To Zwischenlagen - 1
                    Schreiber.WriteStartElement("circle")
                    Schreiber.WriteAttributeString("cx", PunkteZwischenlagen(i).X.ToString(DP))
                    Schreiber.WriteAttributeString("cy", PunkteZwischenlagen(i).Y.ToString(DP))
                    Schreiber.WriteAttributeString("r", PunktRadius.ToString(DP))
                    Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
                    Schreiber.WriteAttributeString("stroke-width", "1")
                    Schreiber.WriteAttributeString("fill", FarbeZwischenlagen)           'SystemColors.ActiveBorderBrush
                    Schreiber.WriteAttributeString("id", "Koppelpunkt")
                    Schreiber.WriteEndElement()  '</circle> - Koppelpunkt der Zwischenlage
                Next
            End If
            'Koppelkurve zeichnen
            Dim KoppelpunkteString As String = KoppelpunkteCollection(0).ToString(DP)  'probieren: KoppelpunkteCollection.ToString(DP) 
            For i As Integer = 1 To AnzPunkteKoppelkurve - 1
                KoppelpunkteString &= (" " & KoppelpunkteCollection(i).ToString(DP))
            Next
            Schreiber.WriteStartElement("polygon")
            Schreiber.WriteAttributeString("points", KoppelpunkteString)
            Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
            Schreiber.WriteAttributeString("stroke-width", "1")
            Schreiber.WriteAttributeString("fill", "none")
            Schreiber.WriteAttributeString("id", "Koppelkurve")
            Schreiber.WriteEndElement()  '</polygon> - Koppelkurve
            'KoppelL1
            Schreiber.WriteStartElement("line")
            Schreiber.WriteAttributeString("x1", KurbelP1.X.ToString(DP))
            Schreiber.WriteAttributeString("y1", (KurbelP1.Y).ToString(DP))
            Schreiber.WriteAttributeString("x2", KoppelP0.X.ToString(DP))
            Schreiber.WriteAttributeString("y2", (KoppelP0.Y).ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeKoppel12)
            Schreiber.WriteAttributeString("stroke-width", "5")
            Schreiber.WriteAttributeString("id", "KoppelL1")
            Schreiber.WriteEndElement()  '</line> - KoppelL1
            'KoppelL2
            Schreiber.WriteStartElement("line")
            Schreiber.WriteAttributeString("x1", KoppelP0.X.ToString(DP))
            Schreiber.WriteAttributeString("y1", (KoppelP0.Y).ToString(DP))
            Schreiber.WriteAttributeString("x2", KoppelP1.X.ToString(DP))
            Schreiber.WriteAttributeString("y2", (KoppelP1.Y).ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeKoppel12)
            Schreiber.WriteAttributeString("stroke-width", "3")
            Schreiber.WriteAttributeString("id", "KoppelL2")
            Schreiber.WriteEndElement()  '</line> - KoppelL2
            'Getriebe=Kurbel-Koppel-Schwinge
            Schreiber.WriteStartElement("polyline")
            Schreiber.WriteAttributeString("points", New Point(KurbelP0.X, KurbelP0.Y).ToString(DP) & " " & New Point(KurbelP1.X, KurbelP1.Y).ToString(DP) & " " & New Point(SchwingeP1.X, SchwingeP1.Y).ToString(DP) & " " & New Point(SchwingeP0.X, SchwingeP0.Y).ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
            Schreiber.WriteAttributeString("stroke-width", "3")
            Schreiber.WriteAttributeString("fill", "none")
            Schreiber.WriteAttributeString("id", "Getriebe")
            Schreiber.WriteEndElement()  '</polyline> - Getriebe=Kurbel-Koppel-Schwinge
            'Kurbel Drehpunkt
            Schreiber.WriteStartElement("circle")
            Schreiber.WriteAttributeString("cx", KurbelP0.X.ToString(DP))
            Schreiber.WriteAttributeString("cy", (KurbelP0.Y).ToString(DP))
            Schreiber.WriteAttributeString("r", PunktRadius.ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
            Schreiber.WriteAttributeString("stroke-width", "1")
            Schreiber.WriteAttributeString("fill", FarbeFüllung)           'SystemColors.ActiveBorderBrush
            Schreiber.WriteAttributeString("id", "KurbelDrehpunkt")
            Schreiber.WriteEndElement()  '</circle> - Kurbel Drehpunkt
            'Schwinge Drehpunkt
            Schreiber.WriteStartElement("circle")
            Schreiber.WriteAttributeString("cx", SchwingeP0.X.ToString(DP))
            Schreiber.WriteAttributeString("cy", (SchwingeP0.Y).ToString(DP))
            Schreiber.WriteAttributeString("r", PunktRadius.ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
            Schreiber.WriteAttributeString("stroke-width", "1")
            Schreiber.WriteAttributeString("fill", FarbeFüllung)           'SystemColors.ActiveBorderBrush
            Schreiber.WriteAttributeString("id", "SchwingeDrehpunkt")
            Schreiber.WriteEndElement()  '</circle> - Schwinge Drehpunkt
            'Koppelpunkt
            Schreiber.WriteStartElement("circle")
            Schreiber.WriteAttributeString("cx", KoppelP1.X.ToString(DP))
            Schreiber.WriteAttributeString("cy", (KoppelP1.Y).ToString(DP))
            Schreiber.WriteAttributeString("r", PunktRadius.ToString(DP))
            Schreiber.WriteAttributeString("stroke", FarbeGetriebe)
            Schreiber.WriteAttributeString("stroke-width", "1")
            Schreiber.WriteAttributeString("fill", FarbeFüllung)           'SystemColors.ActiveBorderBrush
            Schreiber.WriteAttributeString("id", "Koppelpunkt")
            Schreiber.WriteEndElement()  '</circle> - Koppelpunkt
            '--------------------------------------------------------------------------------------------
            ' SVG-Bild Geometrie geschrieben
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteEndElement()   '</g>
            '--------------------------------------------------------------------------------------------
            ' XML-Datei beenden und schließen
            '--------------------------------------------------------------------------------------------
            Schreiber.WriteEndElement()   '</SVG>
            Schreiber.WriteEndDocument()
        End Using

    End Sub

#End Region

#Region " Rückruf"

    Private Sub GetriebeWeiterSchalten()
        '---------------------------------------------------------------------------------------------------------------------------------------
        'Punkte aktualisieren: KurbelP1, SchwingeP1, KoppelP0 und KoppelP1
        '---------------------------------------------------------------------------------------------------------------------------------------
        KurbelWinkel += (AniSchrittWeite * KurbelDrehrichtung)
        If KurbelWinkel > 360 Then KurbelWinkel -= 360
        If KurbelWinkel < 0 Then KurbelWinkel += 360
        '---------------------------------------------------
        'Berechnen: KurbelP1, SchwingeP1, KoppelP0, KoppelP1
        '---------------------------------------------------
        'Dim KurbelBogen As Double = KurbelWinkel * GradBogen
        'Dim Kurbelvektor As Vector = New Vector(Cos(KurbelBogen), Sin(KurbelBogen))
        'Kurbelvektor = Vector.Multiply(Kurbelvektor, KurbelLänge)
        'KurbelP1 = Point.Add(Kurbelvektor, KurbelP0)
        Dim Drehung As New RotateTransform With {.Angle = KurbelWinkel}
        KurbelP1 = KurbelP0 + Drehung.Transform(New Vector(KurbelLänge, 0))
        '------------------------------------------
        'Kurbel in die neue Stellung korrigieren
        '------------------------------------------
        Kurbel.X1 = KurbelP0.X
        Kurbel.Y1 = KurbelP0.Y
        Kurbel.X2 = KurbelP1.X
        Kurbel.Y2 = KurbelP1.Y
        'Folgegrößen Basis
        Dim BasisVektor As Vector = Vector.Subtract(SchwingeP0, KurbelP1)
        Dim BasisLänge As Double = BasisVektor.Length
        Dim BasisLängeQuadrat As Double = BasisVektor.LengthSquared
        Dim BasisBogen As Double = Atan2(BasisVektor.Y, BasisVektor.X)
        'Fehlkonstruktionen nicht zeichnen
        If BasisLänge >= (KoppelLänge + SchwingeLänge) Then Exit Sub
        If KoppelLänge >= (BasisLänge + SchwingeLänge) Then Exit Sub
        If SchwingeLänge >= (BasisLänge + KoppelLänge) Then Exit Sub
        'Grenzlagen behandeln, else allgemeinen Fall
        Dim Bogen As Double
        If BasisLänge = (KoppelLänge + SchwingeLänge) Then          'Strecklage
            Bogen = BasisBogen
        ElseIf KoppelLänge = (BasisLänge + SchwingeLänge) Then      'Strecklage
            Bogen = BasisBogen
        ElseIf SchwingeLänge = (BasisLänge + KoppelLänge) Then      'Strecklage
            Bogen = BasisBogen + Math.PI
        Else                                                        'allgemeiner Fall
            'Cosinussatz 
            Dim CosinusB1 As Double = (KoppelLänge * KoppelLänge + BasisLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * KoppelLänge * BasisLänge)
            Dim DreieckBogen As Double = Math.Acos(CosinusB1)
            Bogen = BasisBogen + DreieckBogen * Montage
        End If
        'Schnittpunkt Koppel und Schwinge -> SchwingeP1
        Dim KoppelEinheitsVektor As Vector = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
        Dim KoppelVektor As Vector = Vector.Multiply(KoppelEinheitsVektor, KoppelLänge)
        SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
        'KoppelP0
        Dim KoppelL1Vektor As Vector = Vector.Multiply(KoppelEinheitsVektor, KoppelL1)
        KoppelP0 = Vector.Add(KoppelL1Vektor, KurbelP1)
        'KoppelP1
        Dim KoppelL2EinheitsVektor As Vector = New Vector(-KoppelEinheitsVektor.Y, KoppelEinheitsVektor.X) 'KoppelEinheitsVektor um 90° gedreht
        Dim KoppelL2Vektor As Vector = Vector.Multiply(KoppelL2EinheitsVektor, KoppelL2)
        KoppelP1 = Point.Add(KoppelL2Vektor, KoppelP0)
        '---------------------------------------------------------------
        'Schwinge und die Koppelglieder in die neue Stellung korrigieren
        '---------------------------------------------------------------
        'Schwinge
        Schwinge.X1 = SchwingeP0.X
        Schwinge.Y1 = SchwingeP0.Y
        Schwinge.X2 = SchwingeP1.X
        Schwinge.Y2 = SchwingeP1.Y
        'Koppel
        Koppel.X1 = KurbelP1.X
        Koppel.Y1 = KurbelP1.Y
        Koppel.X2 = SchwingeP1.X
        Koppel.Y2 = SchwingeP1.Y
        'Koppel1
        Koppel1.X1 = KurbelP1.X
        Koppel1.Y1 = KurbelP1.Y
        Koppel1.X2 = KoppelP0.X
        Koppel1.Y2 = KoppelP0.Y
        'Koppel2
        Koppel2.X1 = KoppelP0.X
        Koppel2.Y1 = KoppelP0.Y
        Koppel2.X2 = KoppelP1.X
        Koppel2.Y2 = KoppelP1.Y
        'Koppel Drehpunkt
        KoppelPunkt.Data = New EllipseGeometry() With {.Center = KoppelP1, .RadiusX = PunktRadius, .RadiusY = PunktRadius}
    End Sub

#End Region

#Region "Hauptfenster-Ereignisse"

    Private Sub Hauptfenster_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If DatenSindVerändert Then
            Dim Meldung As String = "Möchten Sie die Datei" & vbCrLf & LwPfadDatei & vbCrLf
            If IO.File.Exists(LwPfadDatei) Then
                Meldung &= "überschreiben?"
            Else
                Meldung &= "speichern?"
            End If
            Select Case MessageBox.Show(Meldung, "Programm beenden", MessageBoxButton.YesNoCancel)
                Case MessageBoxResult.Yes  ' Ja-Daten speichern, dann beenden
                    XML_SVG_speichern()  'Jetzt speichern
                    e.Cancel = False  'beenden
                Case MessageBoxResult.No ' Nein-Beenden ohne zu speichern
                    e.Cancel = False  'beenden
                    'If Not IsNothing(TabelleDialog) Then TabelleDialog.Close()
                    Transformer = Nothing
            'AktionTimer.Close()
            'AktionTimer.Dispose()
                Case MessageBoxResult.Cancel  'Abbrechen-nicht beenden
                    e.Cancel = True   'halt-nicht beenden
            End Select
        Else  'Programm beenden, Daten sind nicht verändert
            'If Not IsNothing(TabelleDialog) Then TabelleDialog.Close()
            Transformer = Nothing
            'AktionTimer.Close()
            'AktionTimer.Dispose()
        End If
        'Kurbelschwinge StatusXML speichern
        Try
            Dim Settings As New Xml.XmlWriterSettings() With {.Indent = True, .Encoding = System.Text.Encoding.Unicode}
            Using Schreiber As Xml.XmlWriter = Xml.XmlTextWriter.Create(StatusDatei, Settings)
                Schreiber.WriteStartDocument()           '<?xml version="1.0" encoding="UTF-16"?>
                Schreiber.WriteComment("Informationen für das Programm Kurbelschwinge von Wilfried Seibt.")
                Schreiber.WriteStartElement("Kurbelschwinge")
                Schreiber.WriteElementString("Startverzeichnis", IO.Path.GetDirectoryName(LwPfadDatei).ToString)
                Schreiber.WriteEndElement()   '</Kurbelschwinge>
                Schreiber.WriteEndDocument()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Hauptfenster_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Statuszeile.Content = "Fensterbreite=" & Zei.RenderSize.Width.ToString("F", DP) & " Fensterhöhe=" & Zei.RenderSize.Height.ToString("F", DP)
    End Sub

    Private Sub Hauptfenster_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
        If My.Computer.Keyboard.CtrlKeyDown Then
            Select Case e.Key
                Case Key.A : MnZoomAlles(Nothing, Nothing)
                Case Key.B : MnDateiBeenden(Nothing, Nothing)
                Case Key.H : MnAktionBewegen(Nothing, Nothing)
                Case Key.C : MnBeaSchwingeLänge(Nothing, Nothing)
                Case Key.D : MnBeaSchwingeDrehpunkt(Nothing, Nothing)
                Case Key.G : MnZoomGetriebe(Nothing, Nothing)
                Case Key.K : MnBeaKurbelLänge(Nothing, Nothing)
                Case Key.L : MnBeaKoppelLänge(Nothing, Nothing)
                Case Key.M : MnBeaMontage(Nothing, Nothing)
                Case Key.N : MnDateiNeu(Nothing, Nothing)
                Case Key.O : MnDateiOeffnen(Nothing, Nothing)
                Case Key.P : MnBeaKurbelDrehpunkt(Nothing, Nothing)
                Case Key.R : MnBeaKurbelDrehrichtung(Nothing, Nothing)
                Case Key.S : MnDateiSpeichern(Nothing, Nothing)
                Case Key.T : MnBeaTabelle(Nothing, Nothing)
                Case Key.W : MnBeaKurbelWinkel(Nothing, Nothing)
                Case Key.D1 : MnBeaKoppelLänge1(Nothing, Nothing)
                Case Key.D2 : MnBeaKoppelLänge2(Nothing, Nothing)
                Case Else
            End Select
        End If
    End Sub

#End Region

#Region "Zei-Ereignisse"

    Private Sub Zei_Loaded(sender As Object, e As RoutedEventArgs) Handles Zei.Loaded
        'Linke Maustaste Auswahl
        AddHandler Kurbel.PreviewMouseLeftButtonDown, AddressOf Kurbel_PreviewMouseLeftButtonDown
        AddHandler Schwinge.PreviewMouseLeftButtonDown, AddressOf Schwinge_PreviewMouseLeftButtonDown
        AddHandler Koppel.PreviewMouseLeftButtonDown, AddressOf Koppel_PreviewMouseLeftButtonDown

        AddHandler KurbelDrehpunkt.PreviewMouseLeftButtonDown, AddressOf KurbelDrehpunkt_PreviewMouseLeftButtonDown
        AddHandler SchwingeDrehpunkt.PreviewMouseLeftButtonDown, AddressOf SchwingeDrehpunkt_PreviewMouseLeftButtonDown
        AddHandler KoppelPunkt.PreviewMouseLeftButtonDown, AddressOf KoppelPunkt_PreviewMouseLeftButtonDown
        'Rechte Maustaste Auswahl
        AddHandler Kurbel.PreviewMouseRightButtonDown, AddressOf Kurbel_PreviewMouseRightButtonDown
        AddHandler Schwinge.PreviewMouseRightButtonDown, AddressOf Schwinge_PreviewMouseRightButtonDown
        AddHandler Koppel.PreviewMouseRightButtonDown, AddressOf Koppel_PreviewMouseRightButtonDown

        AddHandler KurbelDrehpunkt.PreviewMouseRightButtonDown, AddressOf KurbelDrehpunkt_PreviewMouseRightButtonDown
        AddHandler SchwingeDrehpunkt.PreviewMouseRightButtonDown, AddressOf SchwingeDrehpunkt_PreviewMouseRightButtonDown
        AddHandler KoppelPunkt.PreviewMouseRightButtonDown, AddressOf KoppelPunkt_PreviewMouseRightButtonDown

        ReDim PunkteZwischenlagen(Zwischenlagen - 1)

        'Daten aus der StatusDatei holen
        Dim Startverzeichnis As String = ""

        'Kurbelschwinge StatusDatei laden
        Dim Verzeichnis As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData
        Verzeichnis = Trim(Verzeichnis)
        If Right$(Verzeichnis, 1) <> "\" Then Verzeichnis &= "\"
        StatusDatei = Verzeichnis & "Kurbelschwinge.xml"
        Try
            Dim Settings As New Xml.XmlReaderSettings() With {.DtdProcessing = Xml.DtdProcessing.Parse}
            Using Leser As Xml.XmlReader = Xml.XmlTextReader.Create(StatusDatei, Settings)
                ' Schleife durch alle Elemente, Texte und EndElemente
                Dim StartVerzeichnisLesen As Boolean
                While Leser.Read()
                    Select Case Leser.NodeType
                        Case System.Xml.XmlNodeType.Element
                            If Leser.Name = "Startverzeichnis" Then StartVerzeichnisLesen = True
                        Case System.Xml.XmlNodeType.Text
                            If StartVerzeichnisLesen Then Startverzeichnis = Leser.Value
                        Case System.Xml.XmlNodeType.EndElement
                            If Leser.Name = "Startverzeichnis" Then StartVerzeichnisLesen = False
                    End Select
                End While 'Leser.Read()
            End Using
        Catch ex As Exception
            MessageBox.Show("Fehler beim Lesen der StatusDatei: " & ex.ToString)
        End Try
        If Startverzeichnis = "" Then Startverzeichnis = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Startverzeichnis = Trim(Startverzeichnis)
        If Right$(Startverzeichnis, 1) <> "\" Then Startverzeichnis &= "\"
        LwPfadDatei = Startverzeichnis & "Kurbelschwinge.svg"
        'Startdaten laden und anzeigen
        StartgetriebeZeichnen()
        Zei.Background = PinselFenster
        MnZoomAlles(Nothing, Nothing)
    End Sub

    Private Sub Zei_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Zei.MouseLeftButtonDown
        'Klickpunkt - Koordinaten in mm anzeigen
        Dim MatrixInvertiert As Matrix = Transformer.Matrix
        MatrixInvertiert.Invert()
        Dim KlickMM As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)   'Klickpunkt im mm
        KoordinatenTextBox.Text = KlickMM.X.ToString("F", DP) & "  " & KlickMM.Y.ToString("F", DP)
    End Sub

    Private Sub Zei_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Zei.MouseLeftButtonUp
        If Maus = Aktion.ZoomFenster Then
            'Auswahlrahmen entfernen
            Zei.Children.Remove(Auswahlrahmen)
            'Mauscursor zurücksetzen
            Zei.Cursor = Cursors.Arrow
        End If
        Maus = Aktion.Keine
    End Sub

    Private Sub Zei_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Zei.MouseRightButtonDown
        'nur weiter, wenn keine Mausaktion läuft.
        If Maus = Aktion.Keine Then
            Klickpunkt = e.GetPosition(Zei)
            If Keyboard.Modifiers = ModifierKeys.None Then
                '---------------------------------------------------------------------------
                'Pan starten
                '---------------------------------------------------------------------------
                Maus = Aktion.Pan
                'aktuelle Matrix-Werte lesen
                TransKlickX = Transformer.Matrix.OffsetX
                TransKlickY = Transformer.Matrix.OffsetY
            Else
                '---------------------------------------------------------------------------
                'Zoom Fenster starten
                '---------------------------------------------------------------------------
                Maus = Aktion.ZoomFenster
                Zei.Cursor = Cursors.Cross
                Auswahlrahmen.Width = 1
                Auswahlrahmen.Height = 1
                Canvas.SetLeft(Auswahlrahmen, Klickpunkt.X)
                Canvas.SetTop(Auswahlrahmen, Klickpunkt.Y)
                Zei.Children.Add(Auswahlrahmen)
                '---------------------------------------------------------------------------
            End If
        End If
    End Sub

    Private Sub Zei_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Zei.MouseRightButtonUp
        If Maus = Aktion.ZoomFenster Then
            '--------------------------------------------------------------------------------------------------------------
            ' Zoom Fenster auswerten
            '--------------------------------------------------------------------------------------------------------------
            Dim MatrixInvertiert As Matrix = Transformer.Matrix
            MatrixInvertiert.Invert()
            Dim FensterpunktMM As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)
            Dim KlickpunktMM As Point = Point.Multiply(Klickpunkt, MatrixInvertiert)
            Dim ZoomLiMM As Double = Math.Min(FensterpunktMM.X, KlickpunktMM.X)
            Dim ZoomReMM As Double = Math.Max(FensterpunktMM.X, KlickpunktMM.X)
            Dim ZoomObMM As Double = Math.Max(FensterpunktMM.Y, KlickpunktMM.Y)
            Dim ZoomUnMM As Double = Math.Min(FensterpunktMM.Y, KlickpunktMM.Y)
            'neuen Maßstab nur einstellen, wenn das Fenster groß genug ist.
            Dim DiagonalVektor As Vector = New Point(ZoomLiMM, ZoomUnMM) - New Point(ZoomReMM, ZoomObMM)
            If DiagonalVektor.Length > PunktRadius Then
                'Maßstab bestimmen
                Dim Mx As Double = Zei.ActualWidth / (ZoomReMM - ZoomLiMM)
                Dim My As Double = Zei.ActualHeight / (ZoomObMM - ZoomUnMM)
                Dim TransM, TransX, TransY As Double
                If Mx >= My Then   'Bild im Hochformat
                    TransM = My
                    TransY = TransM * ZoomObMM
                    TransX = (Zei.ActualWidth - TransM * (ZoomLiMM + ZoomReMM)) / 2.0
                Else  'Mx < My   Bild im Breitmormat
                    TransM = Mx
                    TransX = -TransM * ZoomLiMM
                    TransY = (Zei.ActualHeight + TransM * (ZoomUnMM + ZoomObMM)) / 2.0
                End If
                'neue Transformation einstellen
                Transformer.Matrix = New Matrix(TransM, 0, 0, -TransM, TransX, TransY)
            End If
            'Auswahlrahmen entfernen und Mauscursor zurücksetzen
            Zei.Children.Remove(Auswahlrahmen)
            Zei.Cursor = Cursors.Arrow
        End If
        'alle Mausaktionen beenden
        Maus = Aktion.Keine
    End Sub

    Private Sub Zei_MouseMove(sender As Object, e As MouseEventArgs) Handles Zei.MouseMove
        Dim AktPunkt As Point = e.GetPosition(Zei)
        Select Case Maus
            Case Aktion.Pan
                '-------------------------------------------------------------------------------------------------
                'rechte Taste wird gezogen -> Pan bzw. ZoomFenster ausführen
                '-------------------------------------------------------------------------------------------------
                Dim Diff As Point = e.GetPosition(Zei) - Klickpunkt
                Dim TransX As Double = TransKlickX + Diff.X
                Dim TransY As Double = TransKlickY + Diff.Y
                Dim TransM As Double = Transformer.Matrix.M11   'Maßstab unverändert
                Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
            Case Aktion.ZoomFenster
                'Auwahlrahmen aktualisieren
                Fensterpunkt = e.GetPosition(Zei)
                Auswahlrahmen.Width = Math.Max(Math.Abs(Fensterpunkt.X - Klickpunkt.X), 1.0)
                Auswahlrahmen.Height = Math.Max(Math.Abs(Fensterpunkt.Y - Klickpunkt.Y), 1.0)
                Canvas.SetLeft(Auswahlrahmen, Math.Min(Fensterpunkt.X, Klickpunkt.X))
                Canvas.SetTop(Auswahlrahmen, Math.Min(Fensterpunkt.Y, Klickpunkt.Y))
      '-------------------------------------------------------------------------------------------------
      'linke Taste wird gezogen -> ausgewähltes Element bearbeiten
      '-------------------------------------------------------------------------------------------------
            Case Aktion.KurbelDrehpunkt
                Dim MatrixInvertiert As Matrix = Transformer.Matrix
                MatrixInvertiert.Invert()
                Dim KurbelP0Temp As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0Temp, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            Case Aktion.Kurbel
                'Variante 1 mit Division durch Maßstab (Siehe Aktion.Schwinge mit Variante 2)
                Dim KurbelP0Zei As Point = Transformer.Transform(KurbelP0)
                Dim VektorZei As Vector = Vector.Subtract(e.GetPosition(Zei), KurbelP0Zei)
                Dim TransM As Double = Transformer.Matrix.M11
                Dim KurbelLängeTemp As Double = VektorZei.Length / TransM
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLängeTemp, SchwingeP0, SchwingeLänge, KoppelLänge)
            Case Aktion.KurbelWinkel
                Dim KurbelP0Zei As Point = Transformer.Transform(KurbelP0)
                Dim VektorZei As Vector = Vector.Subtract(e.GetPosition(Zei), KurbelP0Zei)
                Dim Vektor0 As New Vector(1, 0)
                KurbelWinkel = Vector.AngleBetween(VektorZei, Vektor0)
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            Case Aktion.SchwingeDrehpunkt
                Dim MatrixInvertiert As Matrix = Transformer.Matrix
                MatrixInvertiert.Invert()
                Dim SchwingeP0Temp As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0Temp, SchwingeLänge, KoppelLänge)
            Case Aktion.Schwinge
                'Variante 2 mit invertierter Matrix (Siehe Aktion.Kurbel mit Variante 1)
                Dim MatrixInvertiert As Matrix = Transformer.Matrix
                MatrixInvertiert.Invert()
                Dim aktPuMM As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)
                Dim VektorMM As Vector = Vector.Subtract(aktPuMM, SchwingeP0)
                Dim SchwingeLängeTemp As Double = VektorMM.Length
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLängeTemp, KoppelLänge)
            Case Aktion.Koppel
                Dim MatrixInvertiert As Matrix = Transformer.Matrix
                MatrixInvertiert.Invert()
                Dim aktPuMM As Point = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)
                Dim VektorMM As Vector = Vector.Subtract(aktPuMM, KurbelP1)
                Dim KoppelLängeTemp As Double = VektorMM.Length
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLängeTemp)
            Case Aktion.KoppelPunkt
                Dim MatrixInvertiert As Matrix = Transformer.Matrix
                MatrixInvertiert.Invert()
                'KopppelP1=aktuelle Mausposition in mm umgerechnet
                KoppelP1 = Point.Multiply(e.GetPosition(Zei), MatrixInvertiert)   'neuer Koppelpunkt
                'KoppelP0 = Lot von KoppelP1 auf die Koppel
                Dim Anstieg As Double = (SchwingeP1.Y - KurbelP1.Y) / (SchwingeP1.X - KurbelP1.X)
                KoppelP0.X = (Anstieg * Anstieg * KurbelP1.X + Anstieg * (KoppelP1.Y - KurbelP1.Y) + KoppelP1.X) / (Anstieg * Anstieg + 1)
                KoppelP0.Y = Anstieg * (KoppelP0.X - KurbelP1.X) + KurbelP1.Y
                'KoppelL1 = Länge des Vektors KoppelP0-KurbelP1
                Dim VektorMM As Vector = Vector.Subtract(KoppelP0, KurbelP1)
                KoppelL1 = VektorMM.Length
                'KoppelL2 = Länge des Vektors KoppelP1-KoppelP0
                VektorMM = Vector.Subtract(KoppelP1, KoppelP0)
                KoppelL2 = VektorMM.Length
                'Vorzeichen von KoppelL1 und KoppelL2 bearbeiten
                If (SchwingeP1.Y - KurbelP1.Y) > (SchwingeP1.X - KurbelP1.X) Then
                    'Koppel steht vertikal
                    If Sign(SchwingeP1.Y - KurbelP1.Y) <> Sign(KoppelP0.Y - KurbelP1.Y) Then
                        KoppelL1 *= -1
                    End If
                    If Sign(SchwingeP1.Y - KurbelP1.Y) = Sign(KoppelP1.X - KoppelP0.X) Then
                        KoppelL2 *= -1
                    End If
                Else
                    'Koppel steht horizontal
                    If Sign(SchwingeP1.X - KurbelP1.X) <> Sign(KoppelP0.X - KurbelP1.X) Then
                        KoppelL1 *= -1
                    End If
                    If Sign(SchwingeP1.X - KurbelP1.X) <> Sign(KoppelP1.Y - KoppelP0.Y) Then
                        KoppelL2 *= -1
                    End If
                End If
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            Case Else
        End Select
        '-------------------------------------------------------------------------------------------------
    End Sub

    Private Sub Zei_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Zei.SizeChanged
        '-----------------------------------------------------------
        'aktuelles Bild auf Fenstermitte schieben, Bildgröße bleibt.
        '-----------------------------------------------------------
        Dim alt As Size = e.PreviousSize
        Dim neu As Size = e.NewSize
        'Transformer-Matrix lesen
        Dim TransM As Double = Transformer.Matrix.M11
        Dim TransU As Double = Transformer.Matrix.OffsetX
        Dim TransV As Double = Transformer.Matrix.OffsetY
        'FensterBreite geändert
        If e.WidthChanged Then TransU += (neu.Width - alt.Width) / 2.0
        'FensterHöhe geändert
        If e.HeightChanged Then TransV += (neu.Height - alt.Height) / 2.0
        'Transformer-Matrix anpassen
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransU, .OffsetY = TransV}
    End Sub

    Private Sub Zei_MouseWheel(sender As Object, e As MouseWheelEventArgs) Handles Zei.MouseWheel
        'Zoomen nur, wenn nichts anderes läuft.
        If Maus = Aktion.Keine Then
            Dim aktPu As Point = e.GetPosition(Zei)
            Dim Zoomfaktor, EinsMinusZF As Double
            If e.Delta > 0 Then
                Zoomfaktor = 1.2
                EinsMinusZF = -0.2
            Else
                Zoomfaktor = 0.8
                EinsMinusZF = 0.2
            End If
            Dim TransM As Double = Transformer.Matrix.M11
            Dim TransX As Double = Transformer.Matrix.OffsetX
            Dim TransY As Double = Transformer.Matrix.OffsetY
            TransM *= Zoomfaktor
            TransX *= Zoomfaktor
            TransY *= Zoomfaktor
            TransX += (aktPu.X * EinsMinusZF)
            TransY += (aktPu.Y * EinsMinusZF)
            Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
        End If

    End Sub

    Private Sub Zei_MouseLeave(sender As Object, e As MouseEventArgs) Handles Zei.MouseLeave
        If Maus = Aktion.ZoomFenster Then
            'Auswahlrahmen entfernen
            Zei.Children.Remove(Auswahlrahmen)
            'Mauscursor zurücksetzen
            Zei.Cursor = Cursors.Arrow
        End If
        Maus = Aktion.Keine
    End Sub

#End Region

#Region "Auswahl - Ereignisse"

    ' Aktion mit linker Maustaste

    Private Sub Kurbel_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        'If (Keyboard.Modifiers And ModifierKeys.Control) > 0 Then -> was ist hier besser?
        If My.Computer.Keyboard.CtrlKeyDown Then
            Maus = Aktion.KurbelWinkel
        Else
            Maus = Aktion.Kurbel
        End If
        'Cursor auf den Anfasser-Punkt KurbelP1 setzen
        Dim KurbelP1Zei As Point = Transformer.Transform(KurbelP1)
        Dim KurbelP1Bildschirm As Point = Zei.PointToScreen(KurbelP1Zei)
        Dim xPosition As Int32 = CInt(KurbelP1Bildschirm.X)
        Dim yPosition As Int32 = CInt(KurbelP1Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    Private Sub Schwinge_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Maus = Aktion.Schwinge
        'Cursor auf den Anfasser-Punkt SchwingeP1 setzen
        Dim SchwingeP1Zei As Point = Transformer.Transform(SchwingeP1)
        Dim SchwingeP1Bildschirm As Point = Zei.PointToScreen(SchwingeP1Zei)
        Dim xPosition As Int32 = CInt(SchwingeP1Bildschirm.X)
        Dim yPosition As Int32 = CInt(SchwingeP1Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    Private Sub Koppel_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Maus = Aktion.Koppel
        'Cursor auf den Anfasser-Punkt SchwingeP1 setzen
        Dim SchwingeP1Zei As Point = Transformer.Transform(SchwingeP1)
        Dim SchwingeP1Bildschirm As Point = Zei.PointToScreen(SchwingeP1Zei)
        Dim xPosition As Int32 = CInt(SchwingeP1Bildschirm.X)
        Dim yPosition As Int32 = CInt(SchwingeP1Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    Private Sub KurbelDrehpunkt_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Maus = Aktion.KurbelDrehpunkt
        'Cursor auf den Anfasser-Punkt KurbelP0 setzen
        Dim KurbelP0Zei As Point = Transformer.Transform(KurbelP0)
        Dim KurbelP0Bildschirm As Point = Zei.PointToScreen(KurbelP0Zei)
        Dim xPosition As Int32 = CInt(KurbelP0Bildschirm.X)
        Dim yPosition As Int32 = CInt(KurbelP0Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    Private Sub SchwingeDrehpunkt_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Maus = Aktion.SchwingeDrehpunkt
        'Cursor auf den Anfasser-Punkt SchwingeP0 setzen
        Dim SchwingeP0Zei As Point = Transformer.Transform(SchwingeP0)
        Dim SchwingeP0Bildschirm As Point = Zei.PointToScreen(SchwingeP0Zei)
        Dim xPosition As Int32 = CInt(SchwingeP0Bildschirm.X)
        Dim yPosition As Int32 = CInt(SchwingeP0Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    Private Sub KoppelPunkt_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Maus = Aktion.KoppelPunkt
        'Cursor auf den Anfasser-Punkt KoppelP1 setzen
        Dim KoppelP1Zei As Point = Transformer.Transform(KoppelP1)
        Dim KoppelP1Bildschirm As Point = Zei.PointToScreen(KoppelP1Zei)
        Dim xPosition As Int32 = CInt(KoppelP1Bildschirm.X)
        Dim yPosition As Int32 = CInt(KoppelP1Bildschirm.Y)
        SetCursorPos(xPosition, yPosition)
    End Sub

    ' Aktion mit rechter Maustaste

    Private Sub Kurbel_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        If My.Computer.Keyboard.CtrlKeyDown Then
            MnBeaKurbelWinkel(Nothing, Nothing)
        Else
            MnBeaKurbelLänge(Nothing, Nothing)
        End If
    End Sub

    Private Sub Schwinge_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        MnBeaSchwingeLänge(Nothing, Nothing)
    End Sub

    Private Sub Koppel_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        MnBeaKoppelLänge(Nothing, Nothing)
    End Sub

    Private Sub KurbelDrehpunkt_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        MnBeaKurbelDrehpunkt(Nothing, Nothing)
    End Sub

    Private Sub SchwingeDrehpunkt_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        MnBeaSchwingeDrehpunkt(Nothing, Nothing)
    End Sub

    Private Sub KoppelPunkt_PreviewMouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        e.Handled = True
        Dim Prompt As String = "Bitte die Längen KoppelL1 und KoppelL2, durch Leerzeichen getrennt, eingeben."
        Prompt &= "Als Dezimaltrennzeichen bitte das Komma verwenden."
        Dim Titel As String = "Lage des Koppelpunktes"
        Dim Istwerte As String = KoppelL1.ToString & Strings.Space(2) & KoppelL2.ToString '(ZahlLeerZahl)
        Dim Eingabe As String = InputBox(Prompt, Titel, Istwerte)
        If Eingabe = String.Empty Then Exit Sub
        Eingabe = Strings.Replace(Eingabe, ",", ".")
        Try
            Dim Punkt As Point = Point.Parse(Eingabe)
            KoppelL1 = Punkt.X
            KoppelL2 = Punkt.Y
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        Catch ex As Exception
            ' Show the exception's message.
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

#Region "Menü Datei"

    Private Sub MnDateiNeu(sender As Object, e As ExecutedRoutedEventArgs)
        Zei.Children.Clear()
        StartgetriebeZeichnen()
        MnZoomAlles(Nothing, Nothing)
    End Sub

    Private Sub MnDateiOeffnen(sender As Object, e As ExecutedRoutedEventArgs)
        '--------------------------------------------------------------------------------------------
        ' SVG-Datei mit Getriebeinformationen laden
        '--------------------------------------------------------------------------------------------
        Dim DateiÖffnenDialog As New Microsoft.Win32.OpenFileDialog With {
        .Filter = "SVG Dateien (*.svg)|*.svg|XML Dateien (*.xml)|*.xml|Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*",
        .InitialDirectory = IO.Path.GetDirectoryName(LwPfadDatei),
        .FileName = "Kurbelschwinge",
        .DefaultExt = "svg"}
        If DateiÖffnenDialog.ShowDialog() = False Then Exit Sub
        Try
            '--------------------------------------------------------------------------------------------
            ' Fehlerbehandlung notwendig wegen vieler Benutzereingaben in der zu lesenden Datei
            '--------------------------------------------------------------------------------------------
            LwPfadDatei = DateiÖffnenDialog.FileName
            Me.Title = LwPfadDatei
            '-----------------------------------------------------------------------------------------------
            ' xmlreader
            '-----------------------------------------------------------------------------------------------
            Zei.Children.Clear()
            Dim Settings As New Xml.XmlReaderSettings() With {.DtdProcessing = Xml.DtdProcessing.Parse}
            Using Leser As Xml.XmlReader = Xml.XmlTextReader.Create(LwPfadDatei, Settings)
                'Schalter anlegen
                Dim xmlObergruppe As String = String.Empty   'titel, desc, g
                Dim xmlGruppe As String = String.Empty
                Dim Erkennung As String = "Kurbelschwinge von Wilfried Seibt"
                ' Schleife durch alle Elemente, Texte und EndElemente
                While Leser.Read()
                    Select Case Leser.NodeType
                        Case System.Xml.XmlNodeType.Element
                            'ein Element gelesen
                            If xmlObergruppe = "desc" Then
                                Select Case xmlGruppe
                                    Case "Getriebe"
                                        Select Case Leser.Name
                                            Case "Kurbel"
                                                While Leser.MoveToNextAttribute()
                                                    Select Case Leser.Name
                                                        Case "Drehpunkt" : KurbelP0 = Point.Parse(Leser.Value)
                                                        Case "Länge" : KurbelLänge = Val(Leser.Value)
                                                        Case "Drehwinkel" : KurbelWinkel = Val(Leser.Value)
                                                        Case "Drehrichtung" : KurbelDrehrichtung = CInt(Leser.Value)
                                                    End Select
                                                End While
                                            Case "Schwinge"
                                                While Leser.MoveToNextAttribute()
                                                    Select Case Leser.Name
                                                        Case "Drehpunkt" : SchwingeP0 = Point.Parse(Leser.Value)
                                                        Case "Länge" : SchwingeLänge = Val(Leser.Value)
                                                    End Select
                                                End While
                                            Case "Koppel"
                                                While Leser.MoveToNextAttribute()
                                                    Select Case Leser.Name
                                                        Case "Länge" : KoppelLänge = Val(Leser.Value)
                                                        Case "Länge1" : KoppelL1 = Val(Leser.Value)
                                                        Case "Länge2" : KoppelL2 = Val(Leser.Value)
                                                        Case "Montage" : Montage = CInt(Leser.Value)
                                                    End Select
                                                End While
                                        End Select
                              'Getriebe - Kennwerte eingelesen
                                    Case "Programm"
                                        Select Case Leser.Name
                                            Case "Optionen"
                                                While Leser.MoveToNextAttribute()
                                                    Select Case Leser.Name
                                                        Case "Strecklagen" : Strecklagen = CBool(Leser.Value) : MenuAnsStrecklagen.IsChecked = Strecklagen
                                                        Case "AnzahlZwischenlagen" : Zwischenlagen = CInt(Leser.Value) : If Zwischenlagen > 0 Then MenuAnsZwischenlagen.IsChecked = True Else MenuAnsZwischenlagen.IsChecked = False
                                                        Case "AnzPunkteKoppelkurve" : AnzPunkteKoppelkurve = CInt(Leser.Value)
                                                    End Select
                                                End While
                                            Case "Aktion"
                                                While Leser.MoveToNextAttribute()
                                                    Select Case Leser.Name
                                                        Case "AniSchrittWeite" : AniSchrittWeite = Val(Leser.Value)
                                                        Case "AniSchrittZeit" : AniSchrittZeit = CInt(Leser.Value)
                                                    End Select
                                                End While
                                        End Select
                                        'Programm - Globalvariable eingelesen
                                    Case Else
                                        '---------------------------------------------------------
                                        'xmlGruppe=String.Empty - evtl. xmlGruppe belegen
                                        '---------------------------------------------------------
                                        Select Case Leser.Name
                                            Case "Getriebe" : xmlGruppe = "Getriebe"
                                            Case "Programm" : xmlGruppe = "Programm"
                                                'Case "Umgebung" : xmlGruppe = "Umgebung"
                                        End Select
                                End Select
                            ElseIf xmlObergruppe = "Umgebung" Then
                                Select Case Leser.Name
                                    Case "line"
                                        Dim Name As String = "Linie"
                                        Dim x1, y1, x2, y2 As Double
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "x1" : x1 = Val(Leser.Value)
                                                Case "y1" : y1 = Val(Leser.Value)
                                                Case "x2" : x2 = Val(Leser.Value)
                                                Case "y2" : y2 = Val(Leser.Value) ': MessageBox.Show("y2=" & y2.ToString)
                                            End Select
                                        End While
                                        ' line als Zei.children eintragen
                                        UmLinie = New Line With {.Name = Name, .X1 = x1, .Y1 = y1, .X2 = x2, .Y2 = y2, .Stroke = PinselUmLinien, .StrokeThickness = 1, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmLinie)
                                    Case "rect"
                                        Dim Name As String = "Rechteck"
                                        'Dim x As String = "0"
                                        'Dim y As String = "0"
                                        'Dim Breite As String = "0"
                                        'Dim Höhe As String = "0"
                                        Dim Füllung As Brush = Nothing
                                        Dim x, y, Breite, Höhe As Double
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "x" : x = Val(Leser.Value)
                                                Case "y" : y = Val(Leser.Value)
                                                Case "width" : Breite = Val(Leser.Value)
                                                Case "height" : Höhe = Val(Leser.Value)
                                                Case "fill" : If Leser.Value <> "none" Then Füllung = PinselUmFlächen
                                            End Select
                                        End While
                                        ' rect als Zei.children eintragen
                                        UmRechteckPfad = New Path With {.Name = Name, .Data = New RectangleGeometry(New Rect(x, y, Breite, Höhe)), .Stroke = PinselUmLinien, .StrokeThickness = 1, .Fill = Füllung, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmRechteckPfad)
                                    Case "circle"
                                        Dim Name As String = "Kreis"
                                        'Dim cx As String = "0"
                                        'Dim cy As String = "0"
                                        'Dim r As String = "0"
                                        Dim Füllung As Brush = Nothing
                                        Dim cx, cy, r As Double
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "cx" : cx = Val(Leser.Value)
                                                Case "cy" : cy = Val(Leser.Value)
                                                Case "r" : r = Val(Leser.Value)
                                                Case "fill" : If Leser.Value <> "none" Then Füllung = PinselUmFlächen
                                            End Select
                                        End While
                                        ' circle als Zei.children eintragen
                                        UmKreisPfad = New Path With {.Name = Name, .Data = New EllipseGeometry() With {.Center = New Point(cx, cy), .RadiusX = r, .RadiusY = r}, .Stroke = PinselUmLinien, .StrokeThickness = 1, .Fill = Füllung, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmKreisPfad)
                                    Case "ellipse"
                                        Dim Name As String = "Ellipse"
                                        'Dim cx As String = "0"
                                        'Dim cy As String = "0"
                                        'Dim rx As String = "0"
                                        'Dim ry As String = "0"
                                        Dim Füllung As Brush = Nothing
                                        Dim cx, cy, rx, ry As Double
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "cx" : cx = Val(Leser.Value)
                                                Case "cy" : cy = Val(Leser.Value)
                                                Case "rx" : rx = Val(Leser.Value)
                                                Case "ry" : ry = Val(Leser.Value)
                                                Case "fill" : If Leser.Value <> "none" Then Füllung = PinselUmFlächen
                                            End Select
                                        End While
                                        ' ellipse als Zei.children eintragen
                                        UmEllipsePfad = New Path With {.Name = Name, .Data = New EllipseGeometry() With {.Center = New Point(cx, cy), .RadiusX = rx, .RadiusY = ry}, .Stroke = PinselUmLinien, .StrokeThickness = 1, .Fill = Füllung, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmEllipsePfad)
                                    Case "polyline"
                                        Dim Name As String = "Polylinie"
                                        Dim PunkteString As String = "0,0"
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "points" : PunkteString = Leser.Value
                                            End Select
                                        End While
                                        ' polyline als Zei.children eintragen
                                        UmPolyline = New Polyline With {.Name = Name, .Points = PointCollection.Parse(PunkteString), .Stroke = PinselUmLinien, .StrokeThickness = 1, .Fill = Nothing, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmPolyline)
                                    Case "polygon"
                                        Dim Name As String = "Polygon"
                                        Dim PunkteString As String = "0,0"
                                        Dim Füllung As Brush = Nothing
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "points" : PunkteString = Leser.Value
                                                Case "fill" : If Leser.Value <> "none" Then Füllung = PinselUmFlächen
                                            End Select
                                        End While
                                        ' polyline als Zei.children eintragen
                                        UmPolygon = New Polygon With {.Name = Name, .Points = PointCollection.Parse(PunkteString), .Stroke = PinselUmLinien, .StrokeThickness = 1, .Fill = Füllung, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmPolygon)
                                    Case "path"
                                        Dim Name As String = "Pfad"
                                        Dim DataString As String = String.Empty
                                        Dim Füllung As Brush = Nothing
                                        While Leser.MoveToNextAttribute()
                                            Select Case Leser.Name
                                                Case "id" : Name = Leser.Value
                                                Case "d" : DataString = Leser.Value
                                                Case "fill" : If Leser.Value <> "none" Then Füllung = PinselUmFlächen
                                            End Select
                                        End While
                                        ' circle als Zei.children eintragen
                                        UmPfadPfad = New Path With {.Name = Name, .Data = New PathGeometry With {.Figures = PathFigureCollection.Parse(DataString)}, .Stroke = PinselUmLinien, .StrokeThickness = 1, .ToolTip = .Name, .RenderTransform = Transformer}
                                        Zei.Children.Add(UmPfadPfad)
                                End Select
                            Else
                                '---------------------------------------------------------
                                'xmlObergruppe=String.Empty - evtl. xmlObergruppe belegen
                                '---------------------------------------------------------
                                If Not Leser.IsEmptyElement Then 'bei Leer-Elementen werden keine Schalter gesetzt
                                    Select Case Leser.Name
                                        Case "titel" : xmlObergruppe = "titel"
                                        Case "desc" : xmlObergruppe = "desc"
                                        Case "g"
                                            If Leser.HasAttributes Then
                                                If Leser.MoveToAttribute("id") Then
                                                    If Leser.Value = "Umgebung" Then xmlObergruppe = "Umgebung"
                                                End If
                                                Leser.MoveToElement() 'Moves the reader back to the element node.
                                            End If
                                            'End If
                                    End Select
                                End If
                            End If
                        Case System.Xml.XmlNodeType.Text
                            ' ein Text gelesen
                            If xmlObergruppe = "titel" Then
                                Dim xmlTitel As String = Leser.Value
                                If xmlTitel <> Erkennung Then
                                    MessageBox.Show("Titel: " & xmlTitel & " passt nicht zu eine Kurbelschwinge-SVG", "Falsche Datei", MessageBoxButton.OK, MessageBoxImage.Stop)
                                    Exit Sub
                                End If
                            End If
                        Case System.Xml.XmlNodeType.EndElement
                            ' ein End-Element gelesen - Schalter rücksetzen
                            Select Case Leser.Name
                                Case "titel" : xmlObergruppe = String.Empty
                                Case "desc" : xmlObergruppe = String.Empty
                                Case "g" : xmlObergruppe = String.Empty    'Ende der Umgebungs-Geometrie
                                Case "Kurbel" : xmlGruppe = String.Empty
                                Case "Getriebe" : xmlGruppe = String.Empty
                                Case "Programm" : xmlGruppe = String.Empty
                            End Select
                    End Select 'Leser.NodeType
                End While
                'alle Knoten gelesen
            End Using 'Leser
            'Feld für die Zwischenlagen aktualisieren
            ReDim PunkteZwischenlagen(Zwischenlagen - 1)
            'jetzt noch das neue Getriebe auf die Umgebung setzen
            GetriebeIndex = Zei.Children.Count
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            '--------------------------------------------------------------------------------------------
            ' Fehlerbehandlung Ende
            '--------------------------------------------------------------------------------------------
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler beim Datei öffnen", MessageBoxButton.OK, MessageBoxImage.Error)
            StartgetriebeZeichnen()
        End Try
        MnZoomAlles(Nothing, Nothing)
    End Sub

    Private Sub MnDateiSpeichern(sender As Object, e As ExecutedRoutedEventArgs)
        'MessageBox.Show("ApplicationCommands.Save aufgerufen!", "MnDateiSpeichern")
        If System.IO.File.Exists(LwPfadDatei) Then
            If MessageBox.Show("Möchten Sie " & LwPfadDatei & " überschreiben ?", "Datei speichern", MessageBoxButton.YesNo, MessageBoxImage.Exclamation) = MessageBoxResult.Yes Then
                Statuszeile.Content = LwPfadDatei & " wird gespeichert."
                XML_SVG_speichern()  'Jetzt speichern
                DatenSindVerändert = False
            Else 'nein gedrückt - nicht überschreiben
                Exit Sub
            End If
        Else
            Statuszeile.Content = LwPfadDatei & " wird gespeichert."
            XML_SVG_speichern()  'Jetzt speichern
            MessageBox.Show(LwPfadDatei & " gespeichert.", "Datei speichern", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            DatenSindVerändert = False
        End If
    End Sub

    Private Sub MnDateiSpeichernUnter(sender As Object, e As ExecutedRoutedEventArgs)
        '---------------------------------------------------------------------
        ' Kurbelschwinge als XML-Datei mit SVG-Bild speichern
        '--------------------------------------------------------------------------------------------
        ' SaveFileDialog fragt nach, ob vorhandene Datei überschrieben werden soll.
        Dim DateiSpeichernDialog As New Microsoft.Win32.SaveFileDialog With {
        .FileName = IO.Path.GetFileNameWithoutExtension(LwPfadDatei),
        .InitialDirectory = IO.Path.GetDirectoryName(LwPfadDatei),
        .Filter = "SVG Dateien (*.svg)|*.svg|XML Dateien (*.xml)|*.xml|Alle Dateien (*.*)|*.*",
        .DefaultExt = "svg",
        .Title = "Kurbelschwinge in SVG-Datei speichern"}

        'Dim PfadVorhanden As Boolean = DateiSpeichernDialog.CheckPathExists
        'MessageBox.Show("My.Computer.FileSystem.SpecialDirectories.MyDocuments = " & My.Computer.FileSystem.SpecialDirectories.MyDocuments)

        ' Show save file dialog box
        Dim result? As Boolean = DateiSpeichernDialog.ShowDialog()
        ' Process save file dialog box results
        If result = True Then
            ' Save document
            LwPfadDatei = DateiSpeichernDialog.FileName
            Me.Title = LwPfadDatei
            'DateiSpeichernDialog.CustomPlaces.Add(LwPfadDatei)
            Statuszeile.Content = LwPfadDatei & " wird gespeichert."
            XML_SVG_speichern()
            DatenSindVerändert = False
        End If
        '---------------------------------------------------------------------
    End Sub

    Private Sub MnDateiAutoCAD(sender As Object, e As RoutedEventArgs)
        '---------------------------------------------------------------------
        ' Kurbelschwinge als ACAD-Script speichern
        '--------------------------------------------------------------------------------------------
        ' SaveFileDialog fragt nach, ob vorhandene Datei überschrieben werden soll.
        Dim DateiSpeichernDialog As New Microsoft.Win32.SaveFileDialog With {
        .FileName = IO.Path.GetFileNameWithoutExtension(LwPfadDatei),
        .InitialDirectory = IO.Path.GetDirectoryName(LwPfadDatei),
        .Filter = "Skript Dateien (*.scr)|*.scr|Text Dateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*",
        .DefaultExt = "scr",
        .Title = "Kurbelschwinge in AutoCAD Skript speichern"}

        'Dim PfadVorhanden As Boolean = DateiSpeichernDialog.CheckPathExists
        'MessageBox.Show("My.Computer.FileSystem.SpecialDirectories.MyDocuments = " & My.Computer.FileSystem.SpecialDirectories.MyDocuments)

        ' Show save file dialog box
        Dim result? As Boolean = DateiSpeichernDialog.ShowDialog()
        ' Process save file dialog box results
        If result = True Then
            '---------------------------------------------------------------------
            ' Save document
            Dim SkriptDatei As String = DateiSpeichernDialog.FileName
            'SkriptDatei = "C:/Wilfried/Temp/Script.SCR"                             'nur für Debug - dann wieder löschen
            Statuszeile.Content = SkriptDatei & " wird gespeichert."
            Using Schreiber As IO.StreamWriter = New IO.StreamWriter(SkriptDatei)
                'ObjektFang ausschalten
                Schreiber.WriteLine("OFANG Aus")
                '----------------------------------------------------------------------------------------------------------------
                For iZeiKind As Integer = 0 To Zei.Children.Count - 1                                         'GetriebeIndex - 1
                    Dim aktShape As System.Windows.Shapes.Shape = Zei.Children(iZeiKind)
                    Select Case aktShape.ToString
                        Case "System.Windows.Shapes.Line"
                            Dim aktLinie As Line = Zei.Children(iZeiKind)
                            Dim P1 As Point = New Point(aktLinie.X1, aktLinie.Y1)
                            Dim P2 As Point = New Point(aktLinie.X2, aktLinie.Y2)
                            Schreiber.WriteLine("_line " & P1.ToString(DP) & " " & P2.ToString(DP) & " ")
                        Case "System.Windows.Shapes.Polyline"
                            Dim aktPolylinie As Polyline = Zei.Children(iZeiKind)
                            Schreiber.WriteLine("_pline " & aktPolylinie.Points.ToString(DP) & " ")
                        Case "System.Windows.Shapes.Polygon"
                            Dim aktPolygon As Polygon = Zei.Children(iZeiKind)
                            Schreiber.WriteLine("_pline " & aktPolygon.Points.ToString(DP) & " s")
                        Case "System.Windows.Shapes.Path"
                            Dim aktPfad As Path = Zei.Children(iZeiKind)
                            Select Case aktPfad.Data.GetType.ToString
                                Case "System.Windows.Media.RectangleGeometry"
                                    Dim aktRechteckPfad As Path = Zei.Children(iZeiKind)
                                    Dim RechteckGeo As RectangleGeometry = aktRechteckPfad.Data
                                    Dim Rechteck As Rect = RechteckGeo.Rect
                                    Dim LinksUnten As String = Rechteck.BottomLeft.ToString(DP)
                                    Dim RechtsOben As String = Rechteck.TopRight.ToString(DP)
                                    Schreiber.WriteLine("_rectangle " & LinksUnten & " " & RechtsOben)
                                Case "System.Windows.Media.EllipseGeometry"     'Kreise und Ellipsen
                                    Dim aktEllipsePfad As Path = Zei.Children(iZeiKind)
                                    Dim EllipseGeo As EllipseGeometry = aktEllipsePfad.Data
                                    If EllipseGeo.RadiusX = EllipseGeo.RadiusY Then
                                        'Kreis eintragen
                                        Dim Pm As Point = New Point(EllipseGeo.Center.X, EllipseGeo.Center.Y)
                                        Dim Ra As Double = EllipseGeo.RadiusX
                                        Schreiber.WriteLine("_circle " & Pm.ToString(DP) & " " & Ra.ToString(DP))
                                    Else
                                        'Ellipse eintragen
                                        Dim PuLi As Point = New Point(EllipseGeo.Center.X - EllipseGeo.RadiusX, EllipseGeo.Center.Y)
                                        Dim PuRe As Point = New Point(EllipseGeo.Center.X + EllipseGeo.RadiusX, EllipseGeo.Center.Y)
                                        Dim PuOb As Point = New Point(EllipseGeo.Center.X, EllipseGeo.Center.Y + EllipseGeo.RadiusY)
                                        Schreiber.WriteLine("_ellipse " & PuLi.ToString(DP) & " " & PuRe.ToString(DP) & " " & PuOb.ToString(DP))
                                    End If
                                Case "System.Windows.Media.PathGeometry"
                                    Dim aktPfadPfad As Path = Zei.Children(iZeiKind)
                                    Dim PfadGeo As PathGeometry = aktPfadPfad.Data
                                    Dim DataString As String = PfadGeo.ToString(DP)
                                    'Schreiber.WriteStartElement("path")
                                    'Schreiber.WriteAttributeString("Name", aktPfadPfad.Name)
                                    'Schreiber.WriteAttributeString("d", DataString)
                                    'Schreiber.WriteAttributeString("fill", If(IsNothing(aktPfadPfad.Fill), "none", FarbeUmFlächen))  'PinselUmFlächen oder "none"
                                    'Schreiber.WriteEndElement()  '</path>
                            End Select
                    End Select
                Next ' iZeiKind
                '----------------------------------------------------------------------------------------------------------------
                'ObjektFang wieder einschalten
                Schreiber.WriteLine("OFANG End,Schn,Zen,Bas")
            End Using   'StreamWriter
        End If
        '---------------------------------------------------------------------

    End Sub

    Private Sub MnDateiBeenden(sender As Object, e As ExecutedRoutedEventArgs)
        'MessageBox.Show("ApplicationCommands.Close aufgerufen!", "MnDateiBeenden")
        Me.Close()
    End Sub

#End Region

#Region "Menü Bearbeiten"

    Private Sub MnBeaTabelle(sender As Object, e As RoutedEventArgs)
        Dim TabelleDialog As New Tabelle
        'Dialogfenster mit den aktuellen Werten füllen
        TabelleDialog.TextBoxKurbelDrehwinkel.Text = KurbelWinkel.ToString
        TabelleDialog.TextBoxKurbelX0.Text = KurbelP0.X.ToString
        TabelleDialog.TextBoxKurbelY0.Text = KurbelP0.Y.ToString
        TabelleDialog.TextBoxKurbelLänge.Text = KurbelLänge.ToString
        TabelleDialog.TextBoxSchwingeX0.Text = SchwingeP0.X.ToString
        TabelleDialog.TextBoxSchwingeY0.Text = SchwingeP0.Y.ToString
        TabelleDialog.TextBoxSchwingeLänge.Text = SchwingeLänge.ToString
        TabelleDialog.TextBoxKoppelL1.Text = KoppelL1.ToString
        TabelleDialog.TextBoxKoppelL2.Text = KoppelL2.ToString
        TabelleDialog.TextBoxKoppelLänge.Text = KoppelLänge.ToString
        If KurbelDrehrichtung = 1 Then
            TabelleDialog.CheckBoxKurbelDrehrichtung.IsChecked = True
            TabelleDialog.CheckBoxKurbelDrehrichtung.Content = "positiv"
        Else
            TabelleDialog.CheckBoxKurbelDrehrichtung.IsChecked = False
            TabelleDialog.CheckBoxKurbelDrehrichtung.Content = "negativ"
        End If
        If Montage = 1 Then
            TabelleDialog.CheckBoxKoppelMontage.IsChecked = True
        Else
            TabelleDialog.CheckBoxKoppelMontage.IsChecked = False
        End If
        ' Show window modally - NOTE: Returns only when window is closed
        Dim DialogResultat As Nullable(Of Boolean) = TabelleDialog.ShowDialog()
        'Dim DialogResultat? As Boolean = TabelleDialog.ShowDialog()   'identisch mit voranstehender Zeile
        If DialogResultat Then
            '----------------------------------------------------------------------------------------------
            ' OK gedrückt -> das in der Tabelle eingegebene Getriebe übernehmen
            '----------------------------------------------------------------------------------------------
            KurbelWinkel = Replace(TabelleDialog.TextBoxKurbelDrehwinkel.Text, ".", ",")
            KurbelP0 = New Point(Replace(TabelleDialog.TextBoxKurbelX0.Text, ".", ","), Replace(TabelleDialog.TextBoxKurbelY0.Text, ".", ","))
            KurbelLänge = Replace(TabelleDialog.TextBoxKurbelLänge.Text, ".", ",")
            SchwingeP0 = New Point(Replace(TabelleDialog.TextBoxSchwingeX0.Text, ".", ","), Replace(TabelleDialog.TextBoxSchwingeY0.Text, ".", ","))
            SchwingeLänge = Replace(TabelleDialog.TextBoxSchwingeLänge.Text, ".", ",")
            KoppelL1 = Replace(TabelleDialog.TextBoxKoppelL1.Text, ".", ",")
            KoppelL2 = Replace(TabelleDialog.TextBoxKoppelL2.Text, ".", ",")
            KoppelLänge = Replace(TabelleDialog.TextBoxKoppelLänge.Text, ".", ",")
            KurbelDrehrichtung = IIf(TabelleDialog.CheckBoxKurbelDrehrichtung.IsChecked, 1, -1)
            Montage = IIf(TabelleDialog.CheckBoxKoppelMontage.IsChecked, 1, -1)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        End If
    End Sub

    Private Sub MnBeaKurbelWinkel(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Drehwinkel der Kurbel: ", "Kurbelschwinge", KurbelWinkel.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            KurbelWinkel = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "Kurbel Drehwinkel", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub MnBeaKurbelDrehrichtung(sender As Object, e As RoutedEventArgs)
        If MessageBox.Show("Möchten Sie die Drehrichtung ändern?", "Drehrichtung der Kurbel", MessageBoxButton.YesNo, MessageBoxImage.Question) Then
            KurbelDrehrichtung *= -1
            DatenSindVerändert = True
        End If
    End Sub

    Private Sub MnBeaKurbelDrehpunkt(sender As Object, e As RoutedEventArgs)
        Dim Titel As String = "Drehpunkt der Kurbel"
        Dim Prompt As String = "Bitte X- und Y-Koordinate des Drehpunktes der Kurbel, durch Leerzeichen getrennt, eingeben."
        Prompt &= "Als Dezimaltrennzeichen bitte das Komma verwenden."
        Dim Istwerte As String = KurbelP0.X.ToString & Strings.Space(2) & KurbelP0.Y.ToString '(ZahlLeerZahl)
        Dim Eingabe As String = InputBox(Prompt, Titel, Istwerte)
        If Eingabe = String.Empty Then Exit Sub
        Eingabe = Strings.Replace(Eingabe, ",", ".")
        Try
            Dim KurbelP0Temp As Point = Point.Parse(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0Temp, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            ' Show the exception's message.
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub MnBeaKurbelLänge(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Länge der Kurbel: ", "Kurbelschwinge", KurbelLänge.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            Dim KurbelLängeTemp As Double = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLängeTemp, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "Kurbel Länge", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub MnBeaSchwingeDrehpunkt(sender As Object, e As RoutedEventArgs)
        Dim Prompt As String = "Bitte X- und Y-Koordinate des Drehpunktes der Schwinge, durch Leerzeichen getrennt, eingeben."
        Prompt &= "Als Dezimaltrennzeichen bitte das Komma verwenden."
        Dim Titel As String = "Drehpunkt der Schwinge"
        Dim Istwerte As String = SchwingeP0.X.ToString & Strings.Space(2) & SchwingeP0.Y.ToString '(ZahlLeerZahl)
        Dim Eingabe As String = InputBox(Prompt, Titel, Istwerte)
        If Eingabe = String.Empty Then Exit Sub
        Eingabe = Strings.Replace(Eingabe, ",", ".")
        Try
            Dim SchwingeP0Temp As Point = Point.Parse(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0Temp, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            ' Show the exception's message.
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub MnBeaSchwingeLänge(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Länge der Schwinge: ", "Kurbelschwinge", SchwingeLänge.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            Dim SchwingeLängeTemp As Double = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLängeTemp, KoppelLänge)
            DatenSindVerändert = True
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "SchwingeLänge", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub MnBeaKoppelLänge1(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Länge der Koppel1: ", "Kurbelschwinge", KoppelL1.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            KoppelL1 = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "Koppel1", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub MnBeaKoppelLänge2(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Länge der Koppel2: ", "Kurbelschwinge", KoppelL2.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            KoppelL2 = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "Koppel2", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

    Private Sub MnBeaMontage(sender As Object, e As RoutedEventArgs)
        Montage *= -1
        AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        DatenSindVerändert = True
    End Sub

    Private Sub MnBeaKoppelLänge(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Länge der Koppel: ", "Kurbelschwinge", KoppelLänge.ToString)
        If Eingabe = String.Empty Then Exit Sub   'InputBox Abbrechen gedrückt
        If IsNumeric(Eingabe) Then
            Eingabe = Replace(Eingabe, ".", ",")
            Dim KoppelLängeTemp As Double = CDbl(Eingabe)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLängeTemp)
            DatenSindVerändert = True
        Else
            MessageBox.Show("Eingabefehler: " & Eingabe & " kann nicht als Zahl ausgewertet werden.", "Koppel Länge", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

#End Region

#Region "Menü Aktion"

    Private Sub MnAktionBewegen(sender As Object, e As RoutedEventArgs)
        If GetriebeDreht Then
            AktionTimer.Stop()
            AktionTimer.Dispose()
            GetriebeDreht = False
            MnAktionBewegenHalt.Header = "_Bewegen"
        Else
            ' Create a timer with a millisecond interval.
            AktionTimer = New System.Timers.Timer(AniSchrittZeit)
            ' Hook up the Elapsed event for the timer. 
            AddHandler AktionTimer.Elapsed, AddressOf AktionTimerEreignis
            AktionTimer.AutoReset = True
            AktionTimer.Enabled = True
            GetriebeDreht = True
            MnAktionBewegenHalt.Header = "_Halt"
        End If

    End Sub

    ' The event handler for the Timer.Elapsed event. 
    Private Sub AktionTimerEreignis(source As Object, e As Timers.ElapsedEventArgs)
        'Aufruf von GetriebeWeiterSchalten() über den Dispatcher notwendig, da Zugriff auf das WPF-UserInterface erfolgt!
        Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal, New DelegateZumUI(AddressOf GetriebeWeiterSchalten))
    End Sub

    Private Sub MnAktion_Drehrichtung(sender As Object, e As RoutedEventArgs)
        KurbelDrehrichtung *= -1
        DatenSindVerändert = True
    End Sub

    Private Sub MnAktion_Schrittweite(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("Schrittweite für die Animation in Grad Kurbelwinkel eingeben: ", "Animation", AniSchrittWeite.ToString)
        If IsNumeric(Eingabe) Then AniSchrittWeite = CDbl(Eingabe)
    End Sub

    Private Sub MnAktion_Schrittzeit(sender As Object, e As RoutedEventArgs)
        Dim Eingabe As String = InputBox("SchrittZeit für die Animation: ", "AniSchrittZeit", AniSchrittZeit.ToString)
        If IsNumeric(Eingabe) Then AniSchrittZeit = Abs(CDbl(Eingabe))
    End Sub

#End Region

#Region "Menü Zoom"

    Private Sub MnZoomAlles(sender As Object, e As RoutedEventArgs)
        BildLi = Double.MaxValue
        BildRe = Double.MinValue
        BildUn = Double.MaxValue
        BildOb = Double.MinValue
        'Dim t As String = ""
        'Schleife durch alle Zei.Children
        For i As Integer = 0 To Zei.Children.Count - 1
            Select Case Zei.Children(i).GetType.ToString
                Case "System.Windows.Shapes.Line"
                    Dim Linie As Line = Zei.Children(i)
                    BildLi = Math.Min(Linie.X1, BildLi)
                    BildRe = Math.Max(Linie.X1, BildRe)
                    BildUn = Math.Min(Linie.Y1, BildUn)
                    BildOb = Math.Max(Linie.Y1, BildOb)
                    BildLi = Math.Min(Linie.X2, BildLi)
                    BildRe = Math.Max(Linie.X2, BildRe)
                    BildUn = Math.Min(Linie.Y2, BildUn)
                    BildOb = Math.Max(Linie.Y2, BildOb)
                Case "System.Windows.Shapes.Polyline"
                    Dim Polylinie As Polyline = Zei.Children(i)
                    For Each Punkt As Point In Polylinie.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Polygon"
                    Dim Vieleck As Polygon = Zei.Children(i)
                    For Each Punkt As Point In Vieleck.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Path"
                    Dim Pfad As Path = Zei.Children(i)
                    Dim Geo As Geometry = Pfad.Data
                    Dim Grenzen As Rect = Geo.Bounds
                    BildLi = Math.Min(Grenzen.Left, BildLi)
                    BildRe = Math.Max(Grenzen.Right, BildRe)
                    BildUn = Math.Min(Grenzen.Top, BildUn)
                    BildOb = Math.Max(Grenzen.Bottom, BildOb)
            End Select
            't &= i & " " & " L=" & BildLi & " R=" & BildRe & " U=" & BildUn & " O=" & BildOb & vbCrLf
        Next
        '---------------------------------------------------------------------
        'Extrempunkte der Kurbel
        '---------------------------------------------------------------------
        BildLi = Math.Min(KurbelP0.X - KurbelLänge, BildLi)
        BildRe = Math.Max(KurbelP0.X + KurbelLänge, BildRe)
        BildUn = Math.Min(KurbelP0.Y - KurbelLänge, BildUn)
        BildOb = Math.Max(KurbelP0.Y + KurbelLänge, BildOb)
        '---------------------------------------------------------------------
        'Strecklagen
        '---------------------------------------------------------------------
        'Abstand der Drehpunkte Kurbel und Schwinge berechnen
        Dim GestellVektor As Vector = Vector.Subtract(SchwingeP0, KurbelP0)
        Dim GestellLänge As Double = GestellVektor.Length
        Dim GestellLängeQuadrat As Double = GestellVektor.LengthSquared
        Dim GestellBogen As Double = Atan2(GestellVektor.Y, GestellVektor.X)
        '----------------------------------------------------------------------------------------------------------------------------
        'Strecklage Aussen, Punkt SchwingeP1 beachten
        '----------------------------------------------------------------------------------------------------------------------------
        Dim Streckglied As Double = KurbelLänge + KoppelLänge
        'Cosinussatz 
        Dim CosinusB1 As Double = (Streckglied * Streckglied + GestellLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * Streckglied * GestellLänge)
        Dim DreieckBogen As Double = Math.Acos(CosinusB1)
        Dim Bogen As Double = GestellBogen + DreieckBogen * Montage
        Strecklage1GradKurbel = Bogen * BogenGrad
        Dim StreckgliedEinheitsVektor As Vector = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
        'Schnittpunkt KurbelP1
        Dim KurbelVektor As Vector = Vector.Multiply(StreckgliedEinheitsVektor, KurbelLänge)
        KurbelP1 = Vector.Add(KurbelVektor, KurbelP0)
        'Schnittpunkt SchwingeP1
        Dim KoppelVektor As Vector = Vector.Multiply(StreckgliedEinheitsVektor, KoppelLänge)
        SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
        BildLi = Math.Min(SchwingeP1.X, BildLi)
        BildRe = Math.Max(SchwingeP1.X, BildRe)
        BildUn = Math.Min(SchwingeP1.Y, BildUn)
        BildOb = Math.Max(SchwingeP1.Y, BildOb)
        Dim BoSchwi1 As Double = Atan2(SchwingeP1.Y - SchwingeP0.Y, SchwingeP1.X - SchwingeP0.X)
        '----------------------------------------------------------------------------------------------------------------------------
        'Strecklage innen, Punkt SchwingeP1 beachten
        '----------------------------------------------------------------------------------------------------------------------------
        Streckglied = KoppelLänge - KurbelLänge
        'Cosinussatz 
        CosinusB1 = (Streckglied * Streckglied + GestellLängeQuadrat - SchwingeLänge * SchwingeLänge) / (2 * Streckglied * GestellLänge)
        DreieckBogen = Math.Acos(CosinusB1)
        Bogen = GestellBogen + DreieckBogen * Montage
        Strecklage2GradKurbel = Bogen * BogenGrad + 180
        If Strecklage2GradKurbel > 360 Then Strecklage2GradKurbel -= 360
        StreckgliedEinheitsVektor = New Vector(Math.Cos(Bogen), Math.Sin(Bogen))
        'Schnittpunkt KurbelP1
        KurbelVektor = Vector.Multiply(-StreckgliedEinheitsVektor, KurbelLänge)
        KurbelP1 = Vector.Add(KurbelVektor, KurbelP0)
        'Schnittpunkt SchwingeP1
        KoppelVektor = Vector.Multiply(StreckgliedEinheitsVektor, KoppelLänge)
        SchwingeP1 = Vector.Add(KoppelVektor, KurbelP1)
        BildLi = Math.Min(SchwingeP1.X, BildLi)
        BildRe = Math.Max(SchwingeP1.X, BildRe)
        BildUn = Math.Min(SchwingeP1.Y, BildUn)
        BildOb = Math.Max(SchwingeP1.Y, BildOb)
        '----------------------------------------------------------------------------------------------------------------------------
        'Horizontal- und Veritikallagen der Schwinge
        '----------------------------------------------------------------------------------------------------------------------------
        Dim BoSchwi2 As Double = Atan2(SchwingeP1.Y - SchwingeP0.Y, SchwingeP1.X - SchwingeP0.X)
        Dim Bo1, Bo2 As Double 'Ziel: (Bo2-Bo1) < Pi   (180°)
        If BoSchwi2 > BoSchwi1 Then
            If (BoSchwi2 - BoSchwi1) < Math.PI Then
                Bo1 = BoSchwi1
                Bo2 = BoSchwi2
            Else  '(BoSchwi2 - BoSchwi1) > Math.PI
                Bo1 = BoSchwi2
                Bo2 = BoSchwi1 + PI + PI
            End If
        Else  'BoSchwi1 > BoSchwi2
            If (BoSchwi1 - BoSchwi2) < Math.PI Then
                Bo1 = BoSchwi2
                Bo2 = BoSchwi1
            Else  '(BoSchwi1 - BoSchwi2) > Math.PI
                Bo1 = BoSchwi1
                Bo2 = BoSchwi2 + PI + PI
            End If
        End If
        'Fünf Extremstellen abfragen und beachten, wenn sie zwischen Bo1 und Bo2 liegen
        If (Bo1 < (-PI / 2)) And (Bo2 > (-PI / 2)) Then BildUn = Math.Min(SchwingeP0.Y - SchwingeLänge, BildUn)
        If (Bo1 < 0) And (Bo2 > 0) Then BildRe = Math.Max(SchwingeP0.X + SchwingeLänge, BildRe)
        If (Bo1 < (PI / 2)) And (Bo2 > (PI / 2)) Then BildOb = Math.Max(SchwingeP0.Y + SchwingeLänge, BildOb)
        If (Bo1 < PI) And (Bo2 > PI) Then BildLi = Math.Min(SchwingeP0.X - SchwingeLänge, BildLi)
        If (Bo1 < (PI / 2 * 3)) And (Bo2 > (PI / 2 * 3)) Then BildUn = Math.Min(SchwingeP0.Y - SchwingeLänge, BildUn)
        'MessageBox.Show(t)
        'Noch einen Rand für den Koppelpunkt
        BildLi -= PunktRadius
        BildRe += PunktRadius
        BildUn -= PunktRadius
        BildOb += PunktRadius

        'MessageBox.Show("BildLi, BildRe, BildUn, BildOb: " & BildLi & "  " & BildRe & "  " & BildUn & "  " & BildOb)
        'Maßstab bestimmen
        Dim Mx As Double = Zei.ActualWidth / (BildRe - BildLi)
        Dim My As Double = Zei.ActualHeight / (BildOb - BildUn)
        Dim TransM, TransX, TransY As Double
        If Mx >= My Then   'Bild im Hochformat
            TransM = My
            TransY = TransM * BildOb  '-TransM * BildOb
            TransX = (Zei.ActualWidth - TransM * (BildLi + BildRe)) / 2.0
        Else  'Mx < My   Bild im Breitformat
            TransM = Mx
            TransX = -TransM * BildLi
            TransY = (Zei.ActualHeight + TransM * (BildUn + BildOb)) / 2.0
        End If
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
        '---------------------------------------------------------------------
        'Liniendicke berechnen
        '---------------------------------------------------------------------
        Pinseldicke = PinseldickeFaktor / TransM
        AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
    End Sub

    Private Sub MnZoomGetriebe(sender As Object, e As RoutedEventArgs)
        BildLi = Double.MaxValue
        BildRe = Double.MinValue
        BildUn = Double.MaxValue
        BildOb = Double.MinValue
        'Dim t As String = ""
        'Schleife durch alle Zei.Children
        For i As Integer = GetriebeIndex To Zei.Children.Count - 1
            Select Case Zei.Children(i).GetType.ToString
                Case "System.Windows.Shapes.Line"
                    Dim Linie As Line = Zei.Children(i)
                    BildLi = Math.Min(Linie.X1, BildLi)
                    BildRe = Math.Max(Linie.X1, BildRe)
                    BildUn = Math.Min(Linie.Y1, BildUn)
                    BildOb = Math.Max(Linie.Y1, BildOb)
                    BildLi = Math.Min(Linie.X2, BildLi)
                    BildRe = Math.Max(Linie.X2, BildRe)
                    BildUn = Math.Min(Linie.Y2, BildUn)
                    BildOb = Math.Max(Linie.Y2, BildOb)
                Case "System.Windows.Shapes.Polyline"
                    Dim Polylinie As Polyline = Zei.Children(i)
                    For Each Punkt As Point In Polylinie.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Polygon"
                    Dim Vieleck As Polygon = Zei.Children(i)
                    For Each Punkt As Point In Vieleck.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Path"
                    Dim Pfad As Path = Zei.Children(i)
                    Dim Geo As Geometry = Pfad.Data
                    Dim Grenzen As Rect = Geo.Bounds
                    BildLi = Math.Min(Grenzen.Left, BildLi)
                    BildRe = Math.Max(Grenzen.Right, BildRe)
                    BildUn = Math.Min(Grenzen.Top, BildUn)
                    BildOb = Math.Max(Grenzen.Bottom, BildOb)
            End Select
            't &= i & " " & " L=" & BildLi & " R=" & BildRe & " U=" & BildUn & " O=" & BildOb & vbCrLf
        Next
        'MessageBox.Show(t)
        'Noch einen Rand für den Koppelpunkt
        BildLi -= PunktRadius
        BildRe += PunktRadius
        BildUn -= PunktRadius
        BildOb += PunktRadius
        'MessageBox.Show("BildLi, BildRe, BildUn, BildOb: " & BildLi & "  " & BildRe & "  " & BildUn & "  " & BildOb)
        'Maßstab bestimmen
        Dim Mx As Double = Zei.ActualWidth / (BildRe - BildLi)
        Dim My As Double = Zei.ActualHeight / (BildOb - BildUn)
        Dim TransM, TransX, TransY As Double
        If Mx >= My Then   'Bild im Hochformat
            TransM = My
            TransY = TransM * BildOb  '-TransM * BildOb
            TransX = (Zei.ActualWidth - TransM * (BildLi + BildRe)) / 2.0
        Else  'Mx < My   Bild im Breitformat
            TransM = Mx
            TransX = -TransM * BildLi
            TransY = (Zei.ActualHeight + TransM * (BildUn + BildOb)) / 2.0
        End If
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
    End Sub

    Private Sub MnZoomKoppelkurve(sender As Object, e As RoutedEventArgs)
        BildLi = Double.MaxValue
        BildRe = Double.MinValue
        BildUn = Double.MaxValue
        BildOb = Double.MinValue
        For Each Punkt As Point In KoppelpunkteCollection
            BildLi = Math.Min(Punkt.X, BildLi)
            BildRe = Math.Max(Punkt.X, BildRe)
            BildUn = Math.Min(Punkt.Y, BildUn)
            BildOb = Math.Max(Punkt.Y, BildOb)
        Next
        'Noch einen Rand für den Koppelpunkt
        BildLi -= PunktRadius
        BildRe += PunktRadius
        BildUn -= PunktRadius
        BildOb += PunktRadius
        'MessageBox.Show("BildLi, BildRe, BildUn, BildOb: " & BildLi & "  " & BildRe & "  " & BildUn & "  " & BildOb)
        'Maßstab bestimmen
        Dim Mx As Double = Zei.ActualWidth / (BildRe - BildLi)
        Dim My As Double = Zei.ActualHeight / (BildOb - BildUn)
        Dim TransM, TransX, TransY As Double
        If Mx >= My Then   'Bild im Hochformat
            TransM = My
            TransY = TransM * BildOb  '-TransM * BildOb
            TransX = (Zei.ActualWidth - TransM * (BildLi + BildRe)) / 2.0
        Else  'Mx < My   Bild im Breitformat
            TransM = Mx
            TransX = -TransM * BildLi
            TransY = (Zei.ActualHeight + TransM * (BildUn + BildOb)) / 2.0
        End If
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
    End Sub

    Private Sub MnZoomPan(sender As Object, e As RoutedEventArgs)
        Dim Meldung As String = "Standpunkt des Betrachters verschieben" & vbCrLf
        Meldung &= "Verschiebung im Format:" & vbCrLf & vbCrLf
        Meldung &= "   X [Leerzeichen] Y   " & vbCrLf & vbCrLf
        Meldung &= "eingeben" & vbCrLf & vbCrLf
        Meldung &= "Hinweis: Pan besser ausführen durch Ziehen mit der rechten Maustaste"
        Dim Titel As String = "Pan numerisch"
        Dim Antwort As String = Strings.Trim(InputBox(Meldung, Titel))
        If Antwort = "" Then Exit Sub
        Try
            Dim Vec As Vector = Vector.Parse(Antwort)       'gewünschte Verschiebung des des Standpunktes in mm
            Vec = Vector.Multiply(Vec, Transformer.Matrix)  'gewünschte Verschiebung des des Standpunktes in Pixel
            Dim TransM As Double = Transformer.Matrix.M11
            Dim TransX As Double = Transformer.Matrix.OffsetX - Vec.X
            Dim TransY As Double = Transformer.Matrix.OffsetY - Vec.Y
            Transformer.Matrix = New Matrix(TransM, 0, 0, -TransM, TransX, TransY)
        Catch ex As Exception
            MessageBox.Show("Eingabefehler:" & vbCrLf & ex.Message, Titel, MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub MnZoomNäher(sender As Object, e As RoutedEventArgs)
        'Heranzoomen zur Detailansicht
        Dim Zoomfaktor As Double = 2
        Dim XY_Faktor As Double = -0.5
        Dim TransM As Double = Transformer.Matrix.M11
        Dim TransX As Double = Transformer.Matrix.OffsetX
        Dim TransY As Double = Transformer.Matrix.OffsetY
        TransM *= Zoomfaktor
        TransX *= Zoomfaktor
        TransY *= Zoomfaktor
        TransX += (Zei.ActualWidth * XY_Faktor)
        TransY += (Zei.ActualHeight * XY_Faktor)
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
    End Sub

    Private Sub MnZoomWeiter(sender As Object, e As RoutedEventArgs)
        'Wegzoomen zur Panoramaansicht
        Dim Zoomfaktor As Double = 0.5
        Dim XY_Faktor As Double = 0.25
        Dim TransM As Double = Transformer.Matrix.M11
        Dim TransX As Double = Transformer.Matrix.OffsetX
        Dim TransY As Double = Transformer.Matrix.OffsetY
        TransM *= Zoomfaktor
        TransX *= Zoomfaktor
        TransY *= Zoomfaktor
        TransX += (Zei.ActualWidth * XY_Faktor)
        TransY += (Zei.ActualHeight * XY_Faktor)
        Transformer.Matrix = New Matrix With {.M11 = TransM, .M22 = -TransM, .OffsetX = TransX, .OffsetY = TransY}
    End Sub

#End Region

#Region "Ansicht"

    Private Sub MnAnsStrecklagen(sender As Object, e As RoutedEventArgs)
        Strecklagen = Not Strecklagen
        AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        DatenSindVerändert = True
        If Strecklagen Then
            MenuAnsStrecklagen.IsChecked = True
        Else
            MenuAnsStrecklagen.IsChecked = False
        End If
    End Sub

    Private Sub MnAnsZwischenlagen(sender As Object, e As RoutedEventArgs)
        Dim Antwort As String = InputBox("Anzahl Zwischenlagen des Getriebes: ", "Zwischenlagen", Zwischenlagen.ToString)
        If Antwort = String.Empty Then Exit Sub      'Abbrechen gedrückt
        Try
            If Zwischenlagen <> CInt(Antwort) Then
                Zwischenlagen = CInt(Antwort)
                Zwischenlagen = Max(Zwischenlagen, 0)  'Zwischenlagen jetzt >=0
                ReDim PunkteZwischenlagen(Zwischenlagen - 1)
                AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
                DatenSindVerändert = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)   ' Show the exception's message.
        End Try
    End Sub

#End Region

#Region "Menü Optionen"

    Private Sub MnOptFarben(sender As Object, e As RoutedEventArgs)
        Select Case e.Source.Content
            Case "Hell"
                PinselFenster = New SolidColorBrush(Colors.White)
                PinselGetriebe = New SolidColorBrush(Colors.Black)
                PinselFüllung = New SolidColorBrush(Colors.Yellow)
                PinselKoppel12 = New SolidColorBrush(Colors.Blue)
                PinselStrecklagen = New SolidColorBrush(Colors.Cyan)
                PinselZwischenlagen = New SolidColorBrush(Colors.DodgerBlue)
                PinselUmFlächen = New SolidColorBrush(Colors.SaddleBrown)
                PinselUmLinien = New SolidColorBrush(Colors.SandyBrown)
            Case "Dunkel"
                PinselFenster = New SolidColorBrush(Colors.Black)
                PinselGetriebe = New SolidColorBrush(Colors.White)
                PinselFüllung = New SolidColorBrush(Colors.Red)
                PinselKoppel12 = New SolidColorBrush(Colors.Yellow)
                PinselStrecklagen = New SolidColorBrush(Colors.Green)
                PinselZwischenlagen = New SolidColorBrush(Colors.GreenYellow)
                PinselUmFlächen = New SolidColorBrush(Colors.SaddleBrown)
                PinselUmLinien = New SolidColorBrush(Colors.SandyBrown)
            Case "System"
                PinselFenster = New SolidColorBrush(SystemColors.WindowColor)
                PinselGetriebe = New SolidColorBrush(SystemColors.WindowTextColor)
                PinselFüllung = New SolidColorBrush(SystemColors.AppWorkspaceColor)
                PinselKoppel12 = New SolidColorBrush(SystemColors.ActiveCaptionColor)
                PinselStrecklagen = New SolidColorBrush(SystemColors.ActiveCaptionColor)
                PinselZwischenlagen = New SolidColorBrush(SystemColors.InactiveCaptionColor)
                PinselUmFlächen = New SolidColorBrush(SystemColors.WindowFrameColor)
                PinselUmLinien = New SolidColorBrush(SystemColors.ControlDarkColor)
        End Select
        Zei.Background = PinselFenster
        KurbelDrehpunkt.Stroke = PinselGetriebe
        KurbelDrehpunkt.Fill = PinselFüllung
        SchwingeDrehpunkt.Stroke = PinselGetriebe
        SchwingeDrehpunkt.Fill = PinselFüllung
        KoppelPunkt.Stroke = PinselGetriebe
        KoppelPunkt.Fill = PinselFüllung
        Kurbel.Stroke = PinselGetriebe
        Schwinge.Stroke = PinselGetriebe
        Koppel.Stroke = PinselGetriebe
        Koppel1.Stroke = PinselKoppel12
        Koppel2.Stroke = PinselKoppel12
        Koppelkurve.Stroke = PinselGetriebe
        AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
    End Sub

    Private Sub MnOptLiniendicke(sender As Object, e As RoutedEventArgs)
        Dim Antwort As String = InputBox("Bitte Liniendicke ( 0.5 bis 3 ) eingeben: ", "Liniendicke", PinseldickeFaktor.ToString)
        If Antwort <> String.Empty Then
            Antwort = Replace(Antwort, ",", ".")
            If IsNumeric(Antwort) Then PinseldickeFaktor = Val(Antwort)
            Pinseldicke = PinseldickeFaktor / Transformer.Matrix.M11
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        End If
    End Sub

    Private Sub MnOptRunden(sender As Object, e As RoutedEventArgs)
        Dim Antwort As String = InputBox("Anzahl Stellen nach dem Komma: ", "Zahlenwerte runden", AnzNachkommastellen.ToString)
        If Antwort = String.Empty Then Exit Sub      'Abbrechen gedrückt
        If (CInt(Antwort) >= 0) And (CInt(Antwort) < 7) Then
            AnzNachkommastellen = CInt(Antwort)
            DP.NumberDecimalDigits = AnzNachkommastellen
        End If
    End Sub

    Private Sub MnOptKoppelkurve(sender As Object, e As RoutedEventArgs)
        Dim Antwort As String = InputBox("Anzahl Punkte auf der Koppelkurve: " & vbCrLf & "20 24 30 36 40 45 60 72 90 120 180 360", "Auflösung der Koppelkurve", AnzPunkteKoppelkurve.ToString)
        If Antwort = String.Empty Then Exit Sub      'Abbrechen gedrückt
        If (CInt(Antwort) = 20) Or (CInt(Antwort) = 24) Or (CInt(Antwort) = 30) Or (CInt(Antwort) = 36) Or (CInt(Antwort) = 40) Or (CInt(Antwort) = 45) Or (CInt(Antwort) = 60) Or (CInt(Antwort) = 72) Or (CInt(Antwort) = 90) Or (CInt(Antwort) = 120) Or (CInt(Antwort) = 180) Or (CInt(Antwort) = 360) Then
            AnzPunkteKoppelkurve = CInt(Antwort)
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
        End If
    End Sub

#End Region

#Region "Menü Umgebung"

    Private Sub MnUmEinfügenLinie(sender As Object, e As RoutedEventArgs)
        'neue Linie anlegen
        UmLinie = New Line With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
        Try
            'Namen eingeben
            Dim Titel As String = "Linie einfügen"
            Dim Meldung As String = "Bitte den Namen der neuen Linie eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Linie"))
            If Name = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmLinie.Name = Name
            UmLinie.ToolTip = Name
            'Anfangspunkt der Linie
            Titel = "Anfangspunkt der neuen Linie"
            Meldung = "Bitte X- und Y-Koordinaten des Anfangspunktes eingeben:" & vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45"
            Dim Eingabe As String = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Dim Punkt1 As Point = Point.Parse(Eingabe)
            UmLinie.X1 = Punkt1.X
            UmLinie.Y1 = Punkt1.Y
            'Endpunkt der Linie
            Titel = "Endpunkt der neuen Linie"
            Meldung = "Bitte X- und Y-Koordinaten des Endpunktes eingeben:" & vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45"
            Eingabe = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            'Eingabe auswerten
            Dim Punkt2 As Point = Point.Parse(Eingabe)
            UmLinie.X2 = Punkt2.X
            UmLinie.Y2 = Punkt2.Y
            If Punkt1 = Punkt2 Then Throw New ArgumentException("Die Linie hat die Länge 0") 'Aufruf der Ausnahme
            'neue Linie einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)   'Alle Zei.Children nach der Umgebung löschen und neu füllen
            Zei.Children.Add(UmLinie)
            GetriebeIndex += 1   'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try
    End Sub

    Private Sub MnUmEinfügenRechteck(sender As Object, e As RoutedEventArgs)
        'neues Rechteck anlegen
        UmRechteckPfad = New Path With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
        Try
            'Namen eingeben
            Dim Titel As String = "Rechteck einfügen."
            Dim Meldung As String = "Bitte den Namen des neuen Rechtecks eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Rechteck"))
            If Name = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmRechteckPfad.Name = Name
            UmRechteckPfad.ToolTip = Name
            'Füllung abfragen
            Meldung = "Soll das Rechteck eine Füllung erhalten?"
            Select Case MessageBox.Show(Meldung, Titel, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                Case MessageBoxResult.Cancel : Exit Sub 'Abbrechen gedrückt
                Case MessageBoxResult.Yes : UmRechteckPfad.Fill = PinselUmFlächen
                Case MessageBoxResult.No : UmRechteckPfad.Fill = Nothing
            End Select
            'Koordinaten des Rechtecks
            Titel = "Koordinaten des neuen Rechtecks eingeben."
            Meldung = "Bitte die Koordinaten des Eckpunktes links, unten, sowie Breite und Höhe des Rechteckes eingeben:" & vbCrLf
            Meldung &= "Zahlen durch Leerzeichen trennen!" & vbCrLf
            Meldung &= "Format: xLinks yUnten Breite Höhe" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45 200 110,5"
            Dim Eingabe As String = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            'Eingabe auswerten
            Dim VierZahlen As Rect = Rect.Parse(Eingabe) ' x1=Links y1=Rechts Width=Unten Height=Oben
            Dim x1 As Double = VierZahlen.X
            Dim y1 As Double = VierZahlen.Y
            Dim Br As Double = VierZahlen.Width
            Dim Ho As Double = VierZahlen.Height
            UmRechteckPfad.Data = New RectangleGeometry(New Rect(x1, y1, Br, Ho))
            'neues Rechteck einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)      'Alle Zei.Children nach der Umgebung löschen
            Zei.Children.Add(UmRechteckPfad)                                                 'neues Rechteck an die UmgebungsObjekte anhängen
            GetriebeIndex += 1                                                               'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)  'Koppelgetriebe wieder einfügen
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try
    End Sub

    Private Sub MnUmEinfügenPolyLinie(sender As Object, e As RoutedEventArgs)
        'Eine neue Polylinie anlegen
        UmPolyline = New Polyline With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = Nothing, .RenderTransform = Transformer}
        Try
            'Namen eingeben
            Dim Titel As String = "Polylinie einfügen"
            Dim Meldung As String = "Bitte den Namen der neuen Polylinie eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Polylinie"))
            If Name = String.Empty Then Exit Sub     'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmPolyline.Name = Name
            UmPolyline.ToolTip = Name
            'Schleife für die Eingabe der Punkte
            Titel = "Punkt eingeben"
            Meldung = vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45" & vbCrLf
            Meldung &= "[e] bzw. [E]  Ende der Punktfolge."
            Dim PunkteString As String = String.Empty
            Dim i As Integer = 1
            'PunkteString in einer Schleife füllen.
            Do
                Try
                    Dim Eingabe As String = InputBox("Bitte X- und Y-Koordinaten des Punktes " & i.ToString & " eingeben:" & Meldung, Titel)
                    Eingabe = Strings.Trim(Eingabe)
                    Eingabe = Strings.Replace(Eingabe, ",", ".")
                    If Eingabe = String.Empty Then      'Abbrechen gedrückt oder aus Versehen Enter gedrückt
                        If Trim(PunkteString) = String.Empty Then Exit Sub   'war doch Abbrechen gedrückt.
                        If MessageBox.Show(PunkteString & vbCrLf & vbCrLf & "Möchten Sie ihre Eingabe verwerfen?", "Polylinie eingeben", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.No Then
                            Continue Do 'weiter eingeben
                        Else
                            Exit Sub    'verwerfen
                        End If
                    End If
                    If Strings.LCase(Strings.Left(Strings.Trim(Eingabe), 1)) = "e" Then Exit Do    'Ende der Punktfolge
                    Dim Punkt As Point = Point.Parse(Eingabe)
                    PunkteString &= " " & Punkt.ToString(DP)  'neuen Punkt anhängen
                    i += 1
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Fehler - Eingabe ungültig", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
                End Try
            Loop
            If i <= 2 Then Throw New ArgumentException("Zu wenig Punkte für eine Polylinie!") 'Aufruf der Ausnahme
            UmPolyline.Points = PointCollection.Parse(PunkteString)
            'neue Polylinie einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)      'Alle Zei.Children nach der Umgebung löschen
            Zei.Children.Add(UmPolyline)                                                     'neue Polylinie an die UmgebungsObjekte anhängen
            GetriebeIndex += 1                                                               'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)  'Koppelgetriebe wieder einfügen
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try
    End Sub

    Private Sub MnUmEinfügenPolygon(sender As Object, e As RoutedEventArgs)
        'Eine neues Polygon anlegen
        UmPolygon = New Polygon With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
        Try
            'Namen eingeben
            Dim Titel As String = "Polygon einfügen"
            Dim Meldung As String = "Bitte den Namen des neuen Polygons eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Polygon"))
            If Name = String.Empty Then Exit Sub     'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmPolygon.Name = Name
            UmPolygon.ToolTip = Name
            'Füllung abfragen
            Meldung = "Soll das Polygon eine Füllung erhalten?"
            Select Case MessageBox.Show(Meldung, Titel, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                Case MessageBoxResult.Cancel : Exit Sub 'Abbrechen gedrückt
                Case MessageBoxResult.Yes : UmPolygon.Fill = PinselUmFlächen
                Case MessageBoxResult.No : UmPolygon.Fill = Nothing
            End Select
            'Schleife für die Eingabe der Punkte
            Titel = "Punkt eingeben"
            Meldung = vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45" & vbCrLf
            Meldung &= "[e] bzw. [E] Ende der Punktfolge."
            Dim PunkteString As String = String.Empty
            Dim i As Integer = 1
            'PunkteString in einer Schleife füllen.
            Do
                Try
                    Dim Eingabe As String = InputBox("Bitte X- und Y-Koordinaten des Punktes " & i.ToString & " eingeben:" & Meldung, Titel)
                    Eingabe = Strings.Trim(Eingabe)
                    Eingabe = Strings.Replace(Eingabe, ",", ".")
                    If Eingabe = String.Empty Then      'Abbrechen gedrückt oder aus Versehen Enter gedrückt
                        If Trim(PunkteString) = String.Empty Then Exit Sub   'war doch Abbrechen gedrückt.
                        If MessageBox.Show(PunkteString & vbCrLf & vbCrLf & "Möchten Sie ihre Eingabe verwerfen?", "Polygon eingeben", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.No Then
                            Continue Do 'weiter eingeben
                        Else
                            Exit Sub    'verwerfen
                        End If
                    End If
                    If Strings.LCase(Strings.Left(Strings.Trim(Eingabe), 1)) = "e" Then Exit Do    'Ende der Punktfolge
                    Dim Punkt As Point = Point.Parse(Eingabe)
                    PunkteString &= " " & Punkt.ToString(DP)  'neuen Punkt anhängen
                    i += 1
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Fehler - Eingabe ungültig", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
                End Try
            Loop
            If i <= 3 Then Throw New ArgumentException("Zu wenig Punkte für ein Polygon!")   'Aufruf der Ausnahme
            UmPolygon.Points = PointCollection.Parse(PunkteString)
            'neues Polygon einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)      'Alle Zei.Children nach der Umgebung löschen
            Zei.Children.Add(UmPolygon)                                                      'neue Polylinie an die UmgebungsObjekte anhängen
            GetriebeIndex += 1                                                               'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)  'Koppelgetriebe wieder einfügen
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try

    End Sub

    Private Sub MnUmEinfügenKreis(sender As Object, e As RoutedEventArgs)
        'neue Linie anlegen
        UmKreisPfad = New Path With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
        Dim Geo As EllipseGeometry = New EllipseGeometry
        Try
            'Namen eingeben
            Dim Titel As String = "Kreis einfügen"
            Dim Meldung As String = "Bitte den Namen des neuen Kreies eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Kreis"))
            If Name = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmKreisPfad.Name = Name
            UmKreisPfad.ToolTip = Name
            'Füllung abfragen
            Meldung = "Soll der Kreis eine Füllung erhalten?"
            Select Case MessageBox.Show(Meldung, Titel, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                Case MessageBoxResult.Cancel : Exit Sub 'Abbrechen gedrückt
                Case MessageBoxResult.Yes : UmKreisPfad.Fill = PinselUmFlächen
                Case MessageBoxResult.No : UmKreisPfad.Fill = Nothing
            End Select
            'Mittelpunkt des Kreises
            Titel = "Mittelpunkt des neuen Kreises"
            Meldung = "Bitte X- und Y-Koordinaten des Mittelpunktes eingeben:" & vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45"
            Dim Eingabe As String = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Dim MittelPu As Point = Point.Parse(Eingabe)
            Geo.Center = MittelPu
            'Radius des Kreises
            Titel = "Radius des neuen Kreises"
            Meldung = "Bitte Radius des Kreises eingeben:"
            Eingabe = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            'Eingabe auswerten
            Dim Ra As Double = Double.Parse(Eingabe)
            If Ra < 1 Then Throw New ArgumentException("Der Radius sollte größer 1 sein!") 'Aufruf der Ausnahme
            Geo.RadiusX = Ra
            Geo.RadiusY = Ra
            UmKreisPfad.Data = Geo
            'neuen Kreis einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)   'Alle Zei.Children nach der Umgebung löschen und neu füllen
            Zei.Children.Add(UmKreisPfad)
            GetriebeIndex += 1   'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try
    End Sub

    Private Sub MnUmEinfügenEllipse(sender As Object, e As RoutedEventArgs)
        'Umgebung EllipsePfad
        'UmEllipsePfad = New Path With {.Name = "Ellipse1", .Data = New EllipseGeometry() With {.Center = New Point(450, 300), .RadiusX = 20, .RadiusY = 30}, .Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .Fill = PinselUmFlächen, .ToolTip = .Name, .RenderTransform = Transformer}
        UmEllipsePfad = New Path With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
        Dim Geo As EllipseGeometry = New EllipseGeometry
        Try
            'Namen eingeben
            Dim Titel As String = "Ellipse einfügen"
            Dim Meldung As String = "Bitte den Namen der neuen Ellipse eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Ellipse"))
            If Name = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmEllipsePfad.Name = Name
            UmEllipsePfad.ToolTip = Name
            'Füllung abfragen
            Meldung = "Soll die Ellipse eine Füllung erhalten?"
            Select Case MessageBox.Show(Meldung, Titel, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                Case MessageBoxResult.Cancel : Exit Sub 'Abbrechen gedrückt
                Case MessageBoxResult.Yes : UmEllipsePfad.Fill = PinselUmFlächen
                Case MessageBoxResult.No : UmEllipsePfad.Fill = Nothing
            End Select
            'Mittelpunkt der Ellipse
            Titel = "Mittelpunkt der neuen Ellipse"
            Meldung = "Bitte X- und Y-Koordinaten des Mittelpunktes eingeben:" & vbCrLf
            Meldung &= "Trennung der Koordinaten durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: X Y" & vbCrLf
            Meldung &= "Beispiel: -12,3 123,45"
            Dim Eingabe As String = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Dim MittelPu As Point = Point.Parse(Eingabe)
            Geo.Center = MittelPu
            'Radius der Ellipse
            Titel = "Radius der neuen Ellipse"
            Meldung = "Bitte RadiusX und RadiusY der Ellipse eingeben:"
            Meldung &= "Trennung der Radien durch Leerzeichen!" & vbCrLf
            Meldung &= "Format: RadiusX RadiusY" & vbCrLf
            Meldung &= "Beispiel: 50 25,5"
            Eingabe = InputBox(Meldung, Titel)
            Eingabe = Strings.Replace(Eingabe, ",", ".")
            Eingabe = Strings.Trim(Eingabe)
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Dim Radien As Point = Point.Parse(Eingabe)
            'Eingabe auswerten
            Dim RaX As Double = Radien.X
            Dim RaY As Double = Radien.Y
            If RaX < 1 Then Throw New ArgumentException("Der RadiusX sollte größer 1 sein!")    'Aufruf der Ausnahme
            If RaY < 1 Then Throw New ArgumentException("Der RadiusY sollte größer 1 sein!")    'Aufruf der Ausnahme
            Geo.RadiusX = RaX
            Geo.RadiusY = RaY
            UmEllipsePfad.Data = Geo
            'neue Ellipse einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)   'Alle Zei.Children nach der Umgebung löschen und neu füllen
            Zei.Children.Add(UmEllipsePfad)
            GetriebeIndex += 1   'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try

    End Sub

    Private Sub MnUmEinfügenPfad(sender As Object, e As RoutedEventArgs)
        Try
            'neuen Pfad anlegen
            UmPfadPfad = New Path With {.Stroke = PinselUmLinien, .StrokeThickness = Pinseldicke, .RenderTransform = Transformer}
            Dim Geo As PathGeometry = New PathGeometry
            Dim Figuren As PathFigureCollection = New PathFigureCollection
            'Namen eingeben
            Dim Titel As String = "Pfad einfügen"
            Dim Meldung As String = "Bitte den Namen des neuen Pfades eingeben:"
            Dim Name As String = Strings.Trim(InputBox(Meldung, Titel, "Pfad"))
            If Name = String.Empty Then Exit Sub    'Abbrechen gedrückt
            Name = Replace(Name, " ", "_")
            UmPfadPfad.Name = Name
            UmPfadPfad.ToolTip = Name
            'Füllung abfragen
            Meldung = "Soll der Pfad eine Füllung erhalten?"
            Select Case MessageBox.Show(Meldung, Titel, MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                Case MessageBoxResult.Cancel : Exit Sub 'Abbrechen gedrückt
                Case MessageBoxResult.Yes : UmPfadPfad.Fill = PinselUmFlächen
                Case MessageBoxResult.No : UmPfadPfad.Fill = Nothing
            End Select
            'Pfad eingeben
            Titel = "Pfad"
            Meldung = "Bitte Pfad eingeben:"
            Dim Eingabe As String = InputBox(Meldung, Titel)
            Eingabe = Strings.Trim(Eingabe)
            If Eingabe = String.Empty Then Exit Sub    'Abbrechen gedrückt
            'Eingabe auswerten
            Figuren = PathFigureCollection.Parse(Eingabe)
            Geo.Figures = Figuren
            UmPfadPfad.Data = Geo
            'neuen Pfad einfügen
            Zei.Children.RemoveRange(GetriebeIndex, Zei.Children.Count - GetriebeIndex)   'Alle Zei.Children nach der Umgebung löschen und neu füllen
            Zei.Children.Add(UmPfadPfad)
            GetriebeIndex += 1   'zeigt wieder auf den ersten freien Platz nach der Umgebung
            AllesBerechnenUndWennUmlauffähigNeuZeichnen(KurbelP0, KurbelLänge, SchwingeP0, SchwingeLänge, KoppelLänge)
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try

    End Sub

    Private Sub MnUmLöschen(sender As Object, e As RoutedEventArgs)
        Try
            Dim t As String = "Index" & "    " & "Name" & vbCrLf & vbCrLf
            For i As Integer = 0 To GetriebeIndex - 1
                Select Case Zei.Children(i).GetType.ToString
                    Case "System.Windows.Shapes.Line"
                        Dim Linie As Line = Zei.Children(i)
                        Name = Linie.Name
                    Case "System.Windows.Shapes.Polyline"
                        Dim Polylinie As Polyline = Zei.Children(i)
                        Name = Polylinie.Name
                    Case "System.Windows.Shapes.Polygon"
                        Dim Vieleck As Polygon = Zei.Children(i)
                        Name = Vieleck.Name
                    Case "System.Windows.Shapes.Path"
                        Dim Pfad As Path = Zei.Children(i)
                        Name = Pfad.Name
                End Select
                t &= i.ToString & Strings.StrDup(6 - Len(i.ToString) * 2, " ") & Name.ToString & vbCrLf
            Next
            t &= vbCrLf & "Geben Sie den Index des zu löschenden Elementes ein:" & vbCrLf & "([A] für alle)"
            Dim Antwort As String = Trim(InputBox(t, "Element löschen"))
            If Antwort = String.Empty Then Exit Sub
            If Strings.UCase(Antwort) = "A" Then
                'alle Umweltobjekte löschen
                Zei.Children.RemoveRange(0, GetriebeIndex)
                GetriebeIndex = 0
                DatenSindVerändert = True
                Exit Sub
            End If
            Dim Index As Integer = Val(Antwort)
            If Index < 0 Then Throw New ArgumentException("Index kleiner Null ist nicht erlaubt!")      'Aufruf der Ausnahme
            If Index >= GetriebeIndex Then Throw New ArgumentException("Index zu groß!")  'Aufruf der Ausnahme
            'wenn bis jetzt noch kein Fehler auftrat, dann das Index-child aus der Aufzählung entfernen
            Zei.Children.RemoveAt(Index)
            GetriebeIndex -= 1
            DatenSindVerändert = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eingabefehler", MessageBoxButton.OK, MessageBoxImage.Error) ' Show the exception's message.
        End Try
    End Sub

#End Region

#Region "Menü Info"

    Private Sub MnInfoKenngrößenGetriebe(sender As Object, e As RoutedEventArgs)
        Dim s As String = "Aktuelle Kenngrößen der Kurbelschwinge"
        s &= vbCrLf & "--------------------------------------------"
        s &= vbCrLf & " KurbelWinkel = " & KurbelWinkel.ToString("F", DP)
        s &= vbCrLf & " KurbelLänge = " & KurbelLänge.ToString("F", DP)
        s &= vbCrLf & " KurbelP0 = " & KurbelP0.X.ToString("F", DP) & "   " & KurbelP0.Y.ToString("F", DP)
        s &= vbCrLf & " KurbelP1 = " & KurbelP1.X.ToString("F", DP) & "   " & KurbelP1.Y.ToString("F", DP)
        s &= vbCrLf & " KurbelDrehrichtung = " & KurbelDrehrichtung.ToString & vbCrLf
        s &= vbCrLf & " SchwingeLänge = " & SchwingeLänge.ToString("F", DP)
        s &= vbCrLf & " SchwingeP0 = " & SchwingeP0.X.ToString("F", DP) & "   " & SchwingeP0.Y.ToString("F", DP)
        s &= vbCrLf & " SchwingeP1 = " & SchwingeP1.X.ToString("F", DP) & "   " & SchwingeP1.Y.ToString("F", DP) & vbCrLf
        s &= vbCrLf & " KoppelLänge = " & KoppelLänge.ToString("F", DP)
        s &= vbCrLf & " KoppelL1  = " & KoppelL1.ToString("F", DP)
        s &= vbCrLf & " KoppelL2 = " & KoppelL2.ToString("F", DP)
        s &= vbCrLf & " KoppelP0 = " & KoppelP0.X.ToString("F", DP) & "   " & KoppelP0.Y.ToString("F", DP)
        s &= vbCrLf & " KoppelP1 = " & KoppelP1.X.ToString("F", DP) & "   " & KoppelP1.Y.ToString("F", DP) & vbCrLf
        s &= vbCrLf & " Montage = " & Montage.ToString & vbCrLf
        s &= vbCrLf & " Strecklage 1 bei Kurbelwinkel: " & If(Strecklagen, Strecklage1GradKurbel.ToString & "°", "nicht angezeigt")
        s &= vbCrLf & " Strecklage 2 bei Kurbelwinkel: " & If(Strecklagen, Strecklage2GradKurbel.ToString & "°", "nicht angezeigt")
        MessageBox.Show(s, LwPfadDatei)
    End Sub

    Private Sub MnInfoKenngrößenProgramm(sender As Object, e As RoutedEventArgs)
        Dim s As String = "Aktuelle Kenngrößen des Programms Kurbelschwinge.exe"
        s &= vbCrLf & Strings.StrDup(40, "~")
        s &= vbCrLf & "Aktuelle Datei: " & LwPfadDatei
        s &= vbCrLf & "StatusDatei: " & StatusDatei
        s &= vbCrLf & "Verzeichnis der Kurbelschwinge.exe: " & My.Application.Info.DirectoryPath.ToString
        s &= vbCrLf & Strings.StrDup(40, "~")
        s &= vbCrLf & "Maßstab = " & Transformer.Matrix.M11.ToString("F", DP) & vbTab & "TransX = " & Transformer.Matrix.OffsetX.ToString("F", DP) & vbTab & "TransY = " & Transformer.Matrix.OffsetY.ToString("F", DP)
        s &= vbCrLf & "Links = " & BildLi.ToString("F", DP) & "  Rechts = " & BildRe.ToString("F", DP) & "  Unten = " & BildUn.ToString("F", DP) & "  Oben = " & BildOb.ToString("F", DP)
        s &= vbCrLf & "Strecklagen anzeigen: " & IIf(Strecklagen, "ja", "nein")
        s &= vbCrLf & "Anzahl Zwischenlagen: " & IIf(Zwischenlagen, Zwischenlagen.ToString, "keine")
        s &= vbCrLf & "Anzahl Umgebungobjekte: " & GetriebeIndex.ToString
        s &= vbCrLf & "Liniendicke: " & Pinseldicke.ToString("F", DP)
        MessageBox.Show(s, "Kurbelschwinge von W. Seibt")
    End Sub

    Private Sub MnInfoZeichnungsInhalt(sender As Object, e As RoutedEventArgs)
        Dim t As String = "Index   Grenzen des Zeichnungsobjektes   Name" & vbCrLf
        For i As Integer = 0 To Zei.Children.Count - 1
            BildLi = Double.MaxValue
            BildRe = Double.MinValue
            BildUn = Double.MaxValue
            BildOb = Double.MinValue
            Dim Name As String = String.Empty
            Select Case Zei.Children(i).GetType.ToString
                Case "System.Windows.Shapes.Line"
                    Dim Linie As Line = Zei.Children(i)
                    Name = Linie.Name
                    BildLi = Math.Min(Linie.X1, BildLi)
                    BildRe = Math.Max(Linie.X1, BildRe)
                    BildUn = Math.Min(Linie.Y1, BildUn)
                    BildOb = Math.Max(Linie.Y1, BildOb)
                    BildLi = Math.Min(Linie.X2, BildLi)
                    BildRe = Math.Max(Linie.X2, BildRe)
                    BildUn = Math.Min(Linie.Y2, BildUn)
                    BildOb = Math.Max(Linie.Y2, BildOb)
                Case "System.Windows.Shapes.Polyline"
                    Dim Polylinie As Polyline = Zei.Children(i)
                    Name = Polylinie.Name
                    For Each Punkt As Point In Polylinie.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Polygon"
                    Dim Vieleck As Polygon = Zei.Children(i)
                    Name = Vieleck.Name
                    For Each Punkt As Point In Vieleck.Points
                        BildLi = Math.Min(Punkt.X, BildLi)
                        BildRe = Math.Max(Punkt.X, BildRe)
                        BildUn = Math.Min(Punkt.Y, BildUn)
                        BildOb = Math.Max(Punkt.Y, BildOb)
                    Next
                Case "System.Windows.Shapes.Path"
                    Dim Pfad As Path = Zei.Children(i)
                    Name = Pfad.Name
                    Dim Geo As Geometry = Pfad.Data
                    Dim Grenzen As Rect = Geo.Bounds
                    BildLi = Math.Min(Grenzen.Left, BildLi)
                    BildRe = Math.Max(Grenzen.Right, BildRe)
                    BildUn = Math.Min(Grenzen.Top, BildUn)
                    BildOb = Math.Max(Grenzen.Bottom, BildOb)
            End Select
            t &= i.ToString & " Li=" & BildLi.ToString("F", DP) & " Re=" & BildRe.ToString("F", DP) & " Un=" & BildUn.ToString("F", DP) & " Ob=" & BildOb.ToString("F", DP) & "   " & Name & vbCrLf
        Next
        MessageBox.Show(t, LwPfadDatei)
    End Sub

    Private Sub MnInfoHilfeOrientierung(sender As Object, e As RoutedEventArgs)
        Dim Meldung As String = "         Bild schieben -> Ziehen mit rechter Maustaste" & vbCrLf
        Meldung &= "Bild größer-kleiner -> Mausrad" & vbCrLf
        Meldung &= "        Zoom Fenster -> [Strg] + Ziehen mit rechter Maustaste" & vbCrLf

        MessageBox.Show(Meldung, "Orientierung im Bild")
    End Sub

    Private Sub MnInfoHilfeBedienung(sender As Object, e As RoutedEventArgs)
        Dim Meldung As String = "Folgende Objekte bearbeiten durch Ziehen mit der linken Maustaste:" & vbCrLf
        Meldung &= "- Kurbel" & vbCrLf
        Meldung &= "- Schwinge" & vbCrLf
        Meldung &= "- Koppel" & vbCrLf
        Meldung &= "- Drehpunkt Kurbel" & vbCrLf
        Meldung &= "- Drehpunkt Schwinge" & vbCrLf
        Meldung &= "- Koppelpunkt" & vbCrLf & vbCrLf
        Meldung &= "Drehen der Kurbel:" & vbCrLf
        Meldung &= "- [Strg]+Kurbel ziehen mit der linken Maustaste" & vbCrLf

        MessageBox.Show(Meldung, "Getriebe bearbeiten")
    End Sub

    Private Sub MnInfoUeber(sender As Object, e As RoutedEventArgs)
        Try
            Dim t As String
            t = "Allgemeine Informationen über das Programm Kurbelschwinge.exe"
            t &= vbCrLf & Strings.StrDup(46, "~")
            t &= vbCrLf & "Verzeichnis: " & My.Application.Info.DirectoryPath.ToString
            t &= vbCrLf & "Beschreibung: " & My.Application.Info.Description.ToString
            t &= vbCrLf & "Version: " & My.Application.Info.Version.ToString
            t &= vbCrLf & My.Application.Info.Copyright.ToString
            MessageBox.Show(t, My.Application.Info.Title.ToString)
        Catch exp As System.Exception
            ' This catch will trap any unexpected error.
            MessageBox.Show(exp.Message, Me.Content, MessageBoxButton.OK, MessageBoxImage.Information)
        End Try
    End Sub

#End Region

End Class
