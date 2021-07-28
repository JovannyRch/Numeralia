

Public Class pna
    Dim aPaterno, aMaterno, nombres, estadoNacimiento As String
    Dim diaNacimiento, anioNacimiento, mesNacimiento, horaNacimiento, minutoNacimiento As Integer
    Dim diaActual, mesActual, anioActual, horaActual, minActual As Integer
    Dim generoValue As String
    Dim meses(12, 2), signo(12), signoChino, sumaDt As String
    Dim estados(32) As String
    Dim signos(12) As String



    Private Sub Btnsalir_Click(sender As Object, e As EventArgs) Handles Btnsalir.Click
        End
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim Now As Date
        Now = System.DateTime.Now

        TextHoraActual.Text = Now.ToString("HH:mm:ss")
    End Sub

    Private Sub pna_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        Dim Now As Date
        Now = System.DateTime.Now
        TextFechaActual.Text = Now.ToString("dd/MM/yyyy")
    End Sub


    Private Sub setearMeses()
        meses(1, 1) = "Enero" : meses(1, 2) = "31"
        meses(2, 1) = "Febrero" : meses(1, 2) = "28"
        meses(3, 1) = "Marzo" : meses(1, 2) = "31"
        meses(4, 1) = "Abril" : meses(1, 2) = "30"
        meses(5, 1) = "Mayo" : meses(1, 2) = "31"
        meses(6, 1) = "Junio" : meses(1, 2) = "30"
        meses(7, 1) = "Julio" : meses(1, 2) = "31"
        meses(8, 1) = "Agosto" : meses(1, 2) = "31"
        meses(9, 1) = "Septiembre" : meses(1, 2) = "30"
        meses(10, 1) = "Octubre" : meses(1, 2) = "31"
        meses(11, 1) = "Noviembre" : meses(1, 2) = "30"
        meses(12, 1) = "Diciembre" : meses(1, 2) = "31"
    End Sub

    Private Sub setearSignos()
        signos(0) = "Rata"
        signos(1) = "Buey"
        signos(2) = "Tigre"
        signos(3) = "Conejo"
        signos(4) = "Dragón"
        signos(5) = "Serpiente"
        signos(6) = "Caballo"
        signos(7) = "Cabra"
        signos(8) = "Mono"
        signos(9) = "Gallo"
        signos(10) = "Perro"
        signos(11) = "Cerdo"
    End Sub


    Private Sub Btnobtener_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Btnobtener.Click
        CalcularTiempoVivido()

    End Sub


    Private Sub CalcularTiempoVivido()

        setearMeses()
        setearSignos()

        Txtamaterno.Text = UCase(Txtamaterno.Text)
        Txtapaterno.Text = UCase(Txtapaterno.Text)
        Txtnombre.Text = UCase(Txtnombre.Text)

        'Obtención de datos del formulario

        aMaterno = Txtamaterno.Text
        aPaterno = Txtapaterno.Text
        nombres = Txtnombre.Text

        diaNacimiento = Val(CmbBoxdia.Text)
        mesNacimiento = Val(CmbBoxmes.Text)
        anioNacimiento = Val(Txtaño.Text)
        horaNacimiento = Val(Txthora.Text)
        minutoNacimiento = Val(Txtmin.Text)

        diaActual = Val(Strings.Left(TextFechaActual.Text, 2))
        mesActual = Val(Strings.Mid(TextFechaActual.Text, 4, 2))
        anioActual = Val(Strings.Right(TextFechaActual.Text, 4))
        horaActual = Val(Strings.Left(TextHoraActual.Text, 2))
        minActual = Val(Strings.Mid(TextHoraActual.Text, 4, 2))

        diaNacimiento = 27
        mesNacimiento = 7
        anioNacimiento = 2021
        horaNacimiento = 0
        minutoNacimiento = 0

        diaActual = 27
        mesActual = 7
        anioActual = 2021
        horaActual = 11
        minActual = 37


        'Fin de la obtencion de datos

        'Validacion de la fecha de nacimiento

        Dim valido As Boolean
        Dim minutosVividos As Integer
        valido = validarFecha(anioNacimiento, mesNacimiento, diaNacimiento)
        If valido = False Then
            MsgBox("Fecha inválida, verifique sus datos", 0 + 32 + 0, "Aviso")
            Return
        End If

        'Fin validacion
        minutosVividos = calcularMinutosVividos(minutoNacimiento, horaNacimiento, diaNacimiento, mesNacimiento, anioNacimiento, minActual, horaActual, diaActual, mesActual, anioActual)
        Console.WriteLine("Minutos vividos " + minutosVividos.ToString())
        Dim anios, meses, semanas, dias, horas, minutos As Integer

        anios = Math.Truncate(minutosVividos / 525600)
        minutosVividos = minutosVividos Mod 525600

        meses = Math.Truncate(minutosVividos / 44640)
        minutosVividos = minutosVividos Mod 44640

        semanas = Math.Truncate(minutosVividos / 10080)
        minutosVividos = minutosVividos Mod 10080

        dias = Math.Truncate(minutosVividos / 1440)
        minutosVividos = minutosVividos Mod 1440

        horas = Math.Truncate(minutosVividos / 60)
        minutosVividos = minutosVividos Mod 60
        minutos = minutosVividos

        lbl_edad.Text = anios.ToString() + " anio(s), " + meses.ToString() + " mes(es), " + semanas.ToString() + " semana(s), " + dias.ToString() + " dia(s), " + horas.ToString() + " hora(s) " + minutos.ToString() + " minutos(s)"


    End Sub



    Private Function calcularMinutosVividos(minutoNac As Integer, horaNac As Integer, diaNac As Integer, mesNac As Integer, anioNac As Integer, minActual As Integer, horaActual As Integer, diaActual As Integer, mesActual As Integer, anioActual As Integer) As Integer
        Dim i As Integer
        Dim resultado As Integer

        resultado = 0


        If anioNacimiento = anioActual And mesNacimiento = mesNac And diaNac = diaActual And horaNac = horaActual Then
            Return minActual - minutoNac
        End If


        Dim h As Integer
        If anioNacimiento = anioActual And mesNacimiento = mesNac And diaNac = diaActual Then
            h = horaActual - horaNac
        Else
            h = horaActual
        End If

        resultado = minActual


        resultado = resultado + (h * 60)

        If anioNacimiento = anioActual And mesNacimiento = mesNac And diaNac = diaActual Then
            Return resultado
        End If

        resultado = resultado + DiasAMinutos(diaActual - 1)

        If anioNacimiento = anioActual And mesNacimiento = mesNac Then
            Return resultado
        End If

        For i = 1 To mesActual - 1
            Console.WriteLine("Mes ac: " + i.ToString())
            resultado = resultado + DiasAMinutos(obtenerDiasMes(i, anioActual))
        Next

        If (anioNacimiento >= anioActual) Then
            Return resultado
        End If

        For i = anioNacimiento + 1 To anioActual - 1
            Console.WriteLine("Anio: " + i.ToString())
            If esAnioBiciesto(i) Then
                resultado = resultado + DiasAMinutos(366)
            Else
                resultado = resultado + DiasAMinutos(365)
            End If
        Next

        For i = mesNac + 1 To 12
            Console.WriteLine("Mes: " + i.ToString())
            resultado = resultado + DiasAMinutos(obtenerDiasMes(i, anioNac))
        Next

        Dim cantidadDiasMesNacimiento As Integer
        cantidadDiasMesNacimiento = obtenerDiasMes(mesNac, anioNac)
        resultado = resultado + DiasAMinutos((cantidadDiasMesNacimiento - diaNacimiento + 1))

        resultado = resultado + (24 - horaNac + 1) * 60
        resultado = resultado + (60 - minutoNac)



        Return resultado
    End Function

    Private Function DiasAMinutos(dias As Integer) As Integer
        Return dias * 24 * 60
    End Function


    Private Function obtenerDiasMes(ByVal Mes As Integer, ByVal Anio As Integer) As Integer
        If esAnioBiciesto(Anio) And Mes = 2 Then
            Return 29
        End If

        If Mes = 2 Then
            Return 28
        End If

        If Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
            Return 30
        End If

        Return 31
    End Function

    Private Function validarFecha(ByVal Anio As Integer, ByVal Mes As Integer, ByVal Dia As Integer) As Boolean


        Dim cantidadDiasMes As Integer
        cantidadDiasMes = obtenerDiasMes(Mes, Anio)
        If Dia <= cantidadDiasMes Then
            Return True
        End If
        Return False
    End Function

    Private Function esAnioBiciesto(ByVal Anio As Integer) As Boolean
        If (Anio Mod 4 = 0) And (Anio Mod 400 = 0 Or Anio Mod 100 <> 0) Then
            Return True
        End If
        Return False
    End Function

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

End Class


