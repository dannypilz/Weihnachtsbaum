Option Strict On
Module Module1

    Sub Main()

        'Anfangstext und die Variabeln zur Ein- und Ausgabe deklarieren
        Console.WriteLine("Dieses Programm zeichnet einen tollen Weihnachtsbaum, um Sie in festliche" & vbCrLf & "Stimmung zu versetzen. Geben sie einfach das Wort 'Baum' oder 'Weihnachtsbaum'" & vbCrLf & "ein und bestättigen sie mit der Enter-Taste. Bei 'Baum' errscheint nur ein" & vbCrLf & "normaler Tannenbaum... bei 'Weihnachtsbaum' ein  toller geschmückter" & vbCrLf & "Weihnachtsbaum.")
        Dim Eingabe As String = Console.ReadLine()
        Dim Ausgabe As String

        'Abfangen von falschen Wörtern
        While Eingabe <> "Baum" And Eingabe <> "Weihnachtsbaum"
            Console.WriteLine("Seien Sie kein Grinch! Bitte geben sie eines der beiden Wörter ein.")
            Eingabe = Console.ReadLine()
        End While

        'If-Abfrage um zu entscheiden welcher Baum gewünscht ist
        If Eingabe = "Baum" Then

            'Baum ausgeben
            Ausgabe = Baum()
            Console.WriteLine(Ausgabe)

            'Abfragen, was der User als nächstes tun möchte
            Console.WriteLine("Sie haben sich für den normalen Baum entschieden... Der Weihnachtsmann ist" & vbCrLf & "traurig. :(")
            Console.WriteLine("Um das Programm zu beenden geben sie bitte 'Grinch' ein und bestättigen sie" & vbCrLf & "wieder mit der Enter-Taste. Falls Sie das Programm neu starten wollen, um den" & vbCrLf & "Baum oder die andere Version des Baumes sich ausgeben zu lassen, geben sie bitte" & vbCrLf & "'Ho Ho Ho' ein.")
            Eingabe = Console.ReadLine()

            'Abfangen von falschen Wörtern
            While Eingabe <> "Grinch" And Eingabe <> "Ho Ho Ho"
                Console.WriteLine("Aber selbstverständlich dürfen sie sich den Baum noch länger anschauen, um in" & vbCrLf & "Weihnachtsstimmung zu kommen. :)")
                Eingabe = Console.ReadLine()
            End While

            'Beenden oder Neu-starten je nach Entscheidung
            If Eingabe = "Ho Ho Ho" Then
                Main()
            Else
                Exit Sub
            End If

        Else 'automatisch der GeschmückteBaum, da die anderen Wörter von While-Schleife abgefangen wurden

            'Weihnachtsbaum ausgeben
            Ausgabe = GeschmückterBaum()
            Console.WriteLine(Ausgabe)

            'Abfragen, was der User als nächstes tun möchte
            Console.WriteLine("Sie haben mit ihrer Auswahl Weihnachten zum leben erweckt! Der Weihnachtsmann" & vbCrLf & "wird sie bestimmt dafür belohnen. :)")
            Console.WriteLine("Um das Programm zu beenden geben sie bitte 'Grinch' ein und bestättigen sie" & vbCrLf & "wieder mit der Enter-Taste. Falls Sie das Programm neu starten wollen, um den" & vbCrLf & "Baum oder die andere Version des Baumes sich ausgeben zu lassen, geben sie bitte" & vbCrLf & "'Ho Ho Ho' ein.")
            Eingabe = Console.ReadLine()

            'Abfangen von falschen Wörtern
            While Eingabe <> "Grinch" And Eingabe <> "Ho Ho Ho"
                Console.WriteLine("Aber selbstverständlich dürfen sie sich den Baum noch länger anschauen, um in" & vbCrLf & "Weihnachtsstimmung zu kommen. :)")
                Eingabe = Console.ReadLine()
            End While

            'Beenden oder Neu-starten je nach Entscheidung
            If Eingabe = "Ho Ho Ho" Then
                Main()
            Else
                Exit Sub
            End If

        End If

    End Sub

    'Funktion um einen einfachen Baum zu zeichnen
    Function Baum() As String
        Dim Text As String = Leerzeichen(9) & "_^_" & vbCrLf &
            Leerzeichen(8) & "<" & Leerzeichen(3) & ">" & vbCrLf &
            Leerzeichen(7) & "<" & Leerzeichen(5) & ">" & vbCrLf &
            Leerzeichen(6) & "<" & Leerzeichen(7) & ">" & vbCrLf &
            Leerzeichen(5) & "<" & Leerzeichen(9) & ">" & vbCrLf &
            Leerzeichen(4) & "<" & Leerzeichen(11) & ">" & vbCrLf &
            Leerzeichen(3) & "<" & Leerzeichen(13) & ">" & vbCrLf &
            Leerzeichen(2) & "<_______________>" & vbCrLf &
            Leerzeichen(9) & "|_|" & vbCrLf
        Return Text
    End Function

    'Funktion um einen "richtigen" Weihnachtsbaum zu zeichnen
    Function GeschmückterBaum() As String
        Dim Text As String = Leerzeichen(10) & "*" & vbCrLf &
            Leerzeichen(9) & "_^_" & vbCrLf &
            Leerzeichen(8) & "<" & Leerzeichen(3) & ">" & vbCrLf &
            Leerzeichen(7) & "<" & Leerzeichen(3) & "o" & Leerzeichen(1) & ">" & vbCrLf &
            Leerzeichen(6) & "<" & Leerzeichen(1) & "J" & Leerzeichen(3) & "J" & Leerzeichen(1) & ">" & vbCrLf &
            Leerzeichen(5) & "<" & Leerzeichen(1) & "o" & Leerzeichen(7) & ">" & vbCrLf &
            Leerzeichen(4) & "<" & Leerzeichen(7) & "o" & Leerzeichen(3) & ">" & vbCrLf &
            Leerzeichen(3) & "<" & Leerzeichen(2) & "o" & Leerzeichen(7) & "J" & Leerzeichen(2) & ">" & vbCrLf &
            Leerzeichen(2) & "<_______________>" & vbCrLf &
            Leerzeichen(9) & "|_|" & vbCrLf
        Return Text
    End Function

    'Hilfsfunktion zum zeichnen
    Function Leerzeichen(Zahl As Integer) As String
        Dim Text As String = ""

        'Leerzeichen sooft anhängen, wie gewünscht
        For i = 1 To Zahl
            Text &= " "
        Next i

        Return Text
    End Function

End Module
