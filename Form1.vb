Imports vb = Microsoft.VisualBasic
Public Class Form1

    Private Sub Convertir_Click(sender As Object, e As EventArgs) Handles Convertir.Click
        On Error Resume Next
        Dim Valor As Integer = CInt(Numero.Value)

        If Valor = 0 Then Resultado.Text = "Introduzca un Número Valido - Input Valid Number" : Resultado.BackColor = Color.Red : Exit Sub
        Resultado.BackColor = Color.White

        Dim TextoNumero As String = CStr(Valor)
        Dim TextoRestante As String = TextoNumero
        Dim NumeroAhora As String = vbNullString
        Dim Cuenta As Integer = Len(TextoNumero)
        Dim Man As Integer = 0
        Dim Respuesta As String = vbNullString
        For Man = 1 To Cuenta
            NumeroAhora = vbNullString
            If Len(TextoRestante) = 0 Then Exit For
            If Len(TextoRestante) = 5 Then
                NumeroAhora = vb.Left(TextoRestante, 1)
                Select Case NumeroAhora
                    Case "1"
                        Respuesta = Respuesta & "XM"
                    Case "2"
                        Respuesta = Respuesta & "IIXM"
                    Case "3"
                        Respuesta = Respuesta & "IIIXM"
                    Case "4"
                        Respuesta = Respuesta & "LIM"
                    Case "5"
                        Respuesta = Respuesta & "LM"
                    Case "6"
                        Respuesta = Respuesta & "ILM"
                    Case "7"
                        Respuesta = Respuesta & "IILM"
                    Case "8"
                        Respuesta = Respuesta & "IIILM"
                    Case "9"
                        Respuesta = Respuesta & "IDM"
                End Select
            End If
            If Len(TextoRestante) = 4 Then
                NumeroAhora = vb.Left(TextoRestante, 1)
                Select Case NumeroAhora
                    Case "1"
                        Respuesta = Respuesta & "M"
                    Case "2"
                        Respuesta = Respuesta & "IIM"
                    Case "3"
                        Respuesta = Respuesta & "IIIM"
                    Case "4"
                        Respuesta = Respuesta & "MI"
                    Case "5"
                        Respuesta = Respuesta & "VM"
                    Case "6"
                        Respuesta = Respuesta & "IVM"
                    Case "7"
                        Respuesta = Respuesta & "IIVM"
                    Case "8"
                        Respuesta = Respuesta & "IIIVM"
                    Case "9"
                        Respuesta = Respuesta & "IXM"
                End Select
            End If
            If Len(TextoRestante) = 3 Then
                NumeroAhora = vb.Left(TextoRestante, 1)
                Select Case NumeroAhora
                    Case "1"
                        Respuesta = Respuesta & "C"
                    Case "2"
                        Respuesta = Respuesta & "IIC"
                    Case "3"
                        Respuesta = Respuesta & "IIIC"
                    Case "4"
                        Respuesta = Respuesta & "DI"
                    Case "5"
                        Respuesta = Respuesta & "D"
                    Case "6"
                        Respuesta = Respuesta & "ID"
                    Case "7"
                        Respuesta = Respuesta & "IID"
                    Case "8"
                        Respuesta = Respuesta & "IIID"
                    Case "9"
                        Respuesta = Respuesta & "IM"
                End Select
            End If
            If Len(TextoRestante) = 2 Then
                NumeroAhora = vb.Left(TextoRestante, 1)
                Select Case NumeroAhora
                    Case "1"
                        Respuesta = Respuesta & "X"
                    Case "2"
                        Respuesta = Respuesta & "IIX"
                    Case "3"
                        Respuesta = Respuesta & "IIIX"
                    Case "4"
                        Respuesta = Respuesta & "LI"
                    Case "5"
                        Respuesta = Respuesta & "L"
                    Case "6"
                        Respuesta = Respuesta & "IL"
                    Case "7"
                        Respuesta = Respuesta & "IIL"
                    Case "8"
                        Respuesta = Respuesta & "IIIL"
                    Case "9"
                        Respuesta = Respuesta & "IC"
                End Select
            End If
            If Len(TextoRestante) = 1 Then
                NumeroAhora = vb.Left(TextoRestante, 1)
                Select Case NumeroAhora
                    Case "1"
                        Respuesta = Respuesta & "I"
                    Case "2"
                        Respuesta = Respuesta & "II"
                    Case "3"
                        Respuesta = Respuesta & "III"
                    Case "4"
                        Respuesta = Respuesta & "VI"
                    Case "5"
                        Respuesta = Respuesta & "V"
                    Case "6"
                        Respuesta = Respuesta & "IV"
                    Case "7"
                        Respuesta = Respuesta & "IIV"
                    Case "8"
                        Respuesta = Respuesta & "IIIV"
                    Case "9"
                        Respuesta = Respuesta & "IX"
                End Select
            End If
            TextoRestante = vb.Right(TextoNumero, Cuenta - Man)
        Next

        Resultado.Text = Respuesta

    End Sub
End Class
