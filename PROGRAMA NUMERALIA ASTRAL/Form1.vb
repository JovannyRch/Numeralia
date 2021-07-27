Public Class pna
    Dim apaterno, amaterno, normbre, año, mes, dia, generoVar, estado, consapat, consamat, consnom As String

    Private Sub Btnobtener_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Btnobtener.Click
        apaterno = Mid(Txtapaterno.Text, 1, 1) & ObtenerPrimerVocal(Txtapaterno.Text)
        amaterno = Mid(Txtamaterno.Text, 1, 1)
        normbre = Mid(Txtnombre.Text, 1, 1)
        año = Mid(Txtaño.Text, 3, 2)
        mes = CmbBoxmes.Text
        dia = CmbBoxdia.Text
        If RdioBtnmasculino.Checked = True Then generoVar = "H"
        If RdioBtnfemenino.Checked = True Then generoVar = "M"
        estado = Mid(CmbBoxlugar.Text, Len(CmbBoxlugar.Text) - 2, 2)
        consapat = ObtenerPrimerConsonante(Txtapaterno.Text)
        consamat = ObtenerPrimerConsonante(Txtamaterno.Text)
        consnom = ObtenerPrimerConsonante(Txtnombre.Text)

    End Sub
    Private Function ObtenerPrimerVocal(ByVal Cadena As String) As String
        Dim letra As String = ""
        For i = 2 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If (letra = "A" Or letra = "E" Or letra = "I" Or letra = "O" Or letra = "U") Then
                Exit For
            End If
        Next i
        Return letra
    End Function
    Private Function ObtenerPrimerConsonante(ByVal Cadena As String) As String
        Dim letra As String = ""
        For i = 2 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If Not (letra = "A" Or letra = "E" Or letra = "I" Or letra = "O" Or letra = "U") Then
                Exit For
            End If
        Next i
        Return letra
    End Function
    Private Sub Txtapaterno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txtapaterno.TextChanged

    End Sub
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBoxDatosPersonales.Enter

    End Sub

    Private Sub fnacimiento_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fnacimiento.Click

    End Sub

    Private Function Mid(ByVal p1 As String) As String
        Throw New NotImplementedException
    End Function

    Private Function Mid(ByVal Cadena As String, ByVal i As Integer, ByVal p3 As Integer) As String
        Throw New NotImplementedException
    End Function

    Private Sub edad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edad.Click

    End Sub

    Private Sub lbl_gestacion_Click(sender As System.Object, e As System.EventArgs) Handles lbl_gestacion.Click

    End Sub
End Class


