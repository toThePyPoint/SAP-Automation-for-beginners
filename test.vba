Sub UruchomSkryptPython()
    Dim scriptPath As String
    Dim command As String

    scriptPath = "C:\Programowanie\99_Moje_projekty\97_BLOG\01_Zautomatyzuj_SAP\integracja_excel_md04.py"

    command = "python " & scriptPath

    ' Run the command using Shell
    Shell command, vbNormalFocus

End Sub



Sub UruchomSkryptPython2(arg)
    Dim scriptPath As String
    Dim command As String

    scriptPath = "C:\Programowanie\99_Moje_projekty\97_BLOG\01_Zautomatyzuj_SAP\integracja_excel_md04.py"

    ' Stworzenie komendy - jak w konsoli CMD
    command = "python " & scriptPath & " " & arg

    ' Wywołanie komendy za pomocą wbudowanej funkcji Shell
    Shell command, vbNormalFocus

End Sub


Sub WywolajSkrypt()
    Dim arg As String

    arg = """" & "MD04 COHV" & """"

    Call UruchomSkryptPython2(arg)

End Sub



    transakcja = Sheets("List1").Range("A1").Value
    arg = """" & transakcja & """"


Sub WywolajSkrypt()
    Dim arg As String
    Dim transakcja As String

    argumenty = Sheets("List1").Range("A1").Value & "," & Sheets("List1").Range("B1").Value & "," & Sheets("List1").Range("C1").Value
    arg = """" & argumenty & """"

    Call UruchomSkryptPython2(arg)

End Sub

Sub UruchomPythonaPrzezExec()
    Dim objShell As Object
    Dim objExec As Object
    Dim command As String
    Dim pythonOutput As String
    Dim exitCode As Integer
    Dim scriptPath As String

    ' Ścieżka do skryptu Python
    scriptPath = "C:\Temp\Kamil\Prywatne\Programowanie\99_Moje_projekty\97_BLOG\01_Zautomatyzuj_SAP\integracja_excel_md04.py"

    arg = Sheets("List1").Range("A1").Value & "," & Sheets("List1").Range("B1").Value & "," & Sheets("List1").Range("C1").Value

    ' Budowanie komendy do uruchomienia Pythona
    command = """" & "python" & """ """ & scriptPath & """ """ & arg & """"

    ' Utworzenie obiektu WScript.Shell
    Set objShell = CreateObject("WScript.Shell")

    ' Uruchomienie skryptu za pomocą metody .Exec
    Set objExec = objShell.Exec(command)

    ' Czekanie na zakończenie procesu
    '    Pętla sprawdza status procesu. Dopóki jest on uruchomiony (status = 0),
    '    VBA czeka.
    Do While objExec.Status = 0
        DoEvents ' Pozwala Excelowi "oddychać" w trakcie czekania
    Loop

    ' Odczytanie całego wyjścia (wszystkiego, co Python "wydrukował")
    pythonOutput = objExec.StdOut.ReadAll

    ' Wyświetlenie przechwyconych komunikatów w oknie dialogowym
    MsgBox "Odebrano następujące komunikaty z Pythona:" & vbCrLf & vbCrLf & pythonOutput


End Sub