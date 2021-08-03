

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        lbl_edad.Text = ""
        lbl_gestacion.Text = ""
        Txtamaterno.Text = ""
        Txtapaterno.Text = ""
        Txtnombre.Text = ""
        CmbBoxdia.Text = "01"
        CmbBoxlugar.Text = "Estado de México(MC)"
        CmbBoxmes.Text = "01"
        Txtaño.Text = "2000"
        RdioBtnfemenino.Checked = True
        RdioBtnmasculino.Checked = False
        lbl_curp.Text = ""
        lbl_rfc.Text = ""
        lbl_griego.Text = ""
        lbl_chino.Text = ""
        lbl_planeta.Text = ""
        lbl_metal.Text = ""
        lbl_flor.Text = ""
        lbl_elemento.Text = ""
        lbl_piedra.Text = ""
        lbl_numastral.Text = ""
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



    Private Sub CalcularTiempoVivido()


        PictureBoxgriego.SizeMode = PictureBoxSizeMode.StretchImage

        Txtamaterno.Text = UCase(Txtamaterno.Text)
        Txtapaterno.Text = UCase(Txtapaterno.Text)
        Txtnombre.Text = UCase(Txtnombre.Text)

        PictureBoxgriego.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBoxchino.SizeMode = PictureBoxSizeMode.StretchImage

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


        If RdioBtnfemenino.Checked Then
            generoValue = "M"
        Else
            generoValue = "H"
        End If

        estadoNacimiento = Strings.Right(CmbBoxlugar.Text, 4).Replace("(", "").Replace(")", "").Trim()

        'Fin de la obtencion de datos

        'Validacion de la fecha de nacimiento

        Dim valido As Boolean

        valido = validarFecha(anioNacimiento, mesNacimiento, diaNacimiento)
        If valido = False Then
            MsgBox("Fecha inválida, verifique sus datos", 0 + 32 + 0, "Aviso")
            Return
        End If

        'Fin validacion

        calcularEdad(minutoNacimiento, horaNacimiento, diaNacimiento, mesNacimiento, anioNacimiento, minActual, horaActual, diaActual, mesActual, anioActual)

        calcularFechaGestacion(diaNacimiento, mesNacimiento, anioNacimiento)

        obtenerSignos(mesNacimiento, diaNacimiento)

        calcularNumeroAstral(diaNacimiento.ToString(), mesNacimiento.ToString(), anioNacimiento.ToString())

        obtenerSignoChino(anioNacimiento)


        calcularCurp(nombres, aPaterno, aMaterno, anioNacimiento.ToString(), mesNacimiento.ToString(), diaNacimiento.ToString(), generoValue, estadoNacimiento)

    End Sub

    Private Sub calcularCurp(nombres As String, paterno As String, materno As String, anio As String, mes As String, dia As String, genero As String, estado As String)

        Dim inicialPaterno, vocalPaterno As String
        Dim inicialMaterno, inicialNombre As String
        Dim ultimos2AnioNac, mes2Dig, dia2Dig As String
        Dim curp, rfc As String
        Dim primeraConsonateNoInicialAp As String
        Dim primeraConsonateNoInicialAm As String
        Dim primeraConsonateNoInicialNom As String

        If nombres.Contains("MARÍA ") Then
            nombres = nombres.Replace("MARÍA ", "")
        End If

        If nombres.Contains("JOSÉ ") Then
            nombres = nombres.Replace("JOSÉ ", "")
        End If


        inicialPaterno = normalizarLetra(Strings.GetChar(paterno, 1))
        vocalPaterno = normalizarLetra(ObtenerPrimerVocalNoInicial(paterno))
        inicialMaterno = normalizarLetra(Strings.GetChar(materno, 1))
        inicialNombre = normalizarLetra(Strings.GetChar(nombres, 1))
        ultimos2AnioNac = Strings.Right(anio, 2)
        mes2Dig = format2Digitos(mes)
        dia2Dig = format2Digitos(dia)
        primeraConsonateNoInicialAp = normalizarLetra(ObtenerPrimerConsonanteNoInicial(paterno))
        primeraConsonateNoInicialAm = normalizarLetra(ObtenerPrimerConsonanteNoInicial(materno))
        primeraConsonateNoInicialNom = normalizarLetra(ObtenerPrimerConsonanteNoInicial(nombres))

        curp = inicialPaterno + vocalPaterno + inicialMaterno + inicialNombre + ultimos2AnioNac + mes2Dig + dia2Dig + genero + estado + primeraConsonateNoInicialAp + primeraConsonateNoInicialAm + primeraConsonateNoInicialNom + "-XX"

        rfc = inicialPaterno + vocalPaterno + inicialMaterno + inicialNombre + ultimos2AnioNac + mes2Dig + dia2Dig + "-ZZZ"

        lbl_curp.Text = curp
        lbl_rfc.Text = rfc
    End Sub

    Private Function format2Digitos(cadena As String) As String
        If Val(cadena) < 10 Then
            Return "0" + cadena
        Else
            Return cadena
        End If
    End Function


    Private Function normalizarLetra(letra As String) As String
        If letra = "Á" Or letra = "Ä" Then
            Return "A"
        End If

        If letra = "É" Or letra = "Ë" Then
            Return "E"
        End If

        If letra = "Í" Or letra = "Ï" Then
            Return "I"
        End If
        If letra = "Ó" Or letra = "Ö" Then
            Return "O"
        End If

        If letra = "Ú" Or letra = "Ü" Then
            Return "Ú"
        End If
        Return letra
    End Function


    Private Sub obtenerSignos(mes As Integer, dia As Integer)
        If (mes = 3 And dia >= 21) Or (mes = 4 And dia <= 19) Then
            setLabels("Aries", "Fuego", "Jaspe", "Roble", "Rojo", "Marte", "Hierro")
            PictureBoxgriego.Image = My.Resources.aries
        ElseIf (mes = 4 And dia >= 20) Or (mes = 5 And dia <= 20) Then
            setLabels("Tauro", "Tierra", "Cuarzo Rosa", "Eucalipto", "Verde", "Venus", "Cobre")
            PictureBoxgriego.Image = My.Resources.tauro
        ElseIf (mes = 5 And dia >= 21) Or (mes = 6 And dia <= 20) Then
            setLabels("Géminis", "Aire", "Ojo de tigre", "Olivo", "Amarillo", "Mercurio", "Azogue")
            PictureBoxgriego.Image = My.Resources.geminis
        ElseIf (mes = 6 And dia >= 21) Or (mes = 7 And dia <= 22) Then
            setLabels("Cáncer", "Agua", "Perla", "Rosa Blanca", "Amarillo", "Luna", "Plata")
            PictureBoxgriego.Image = My.Resources.cancer
        ElseIf (mes = 7 And dia >= 23) Or (mes = 8 And dia <= 22) Then
            setLabels("Leo", "Fuego", "Diamante", "Amapola", "Rojo", "Sol", "Oro")
            PictureBoxgriego.Image = My.Resources.leo
        ElseIf (mes = 8 And dia >= 23) Or (mes = 9 And dia <= 22) Then
            setLabels("Virgo", "Tierra", "Esmeralda", "Gloria de la Mañana", "Dorado Amarillo", "Mercurio", "Azogue")
            PictureBoxgriego.Image = My.Resources.virgo
        ElseIf (mes = 9 And dia >= 23) Or (mes = 10 And dia <= 22) Then
            setLabels("Libra", "Aire", "Agáta gris", "Ciruelo", "Rosa", "Venus", "Cobre")
            PictureBoxgriego.Image = My.Resources.libra
        ElseIf (mes = 10 And dia >= 23) Or (mes = 11 And dia <= 22) Then
            setLabels("Escorpio", "Agua", "Esmeralda", "Higuera", "Rojo", "Marte y Plúton", "Jaspe sardo")
            PictureBoxgriego.Image = My.Resources.escorpio
        ElseIf (mes = 11 And dia >= 22) Or (mes = 12 And dia <= 21) Then
            setLabels("Sagitario", "Fuego", "Zafiro Azul", "Hortensia", "Azul", "Júpiter", "Estaño")
            PictureBoxgriego.Image = My.Resources.sagitario
        ElseIf (mes = 12 And dia >= 22) Or (mes = 1 And dia <= 19) Then
            setLabels("Capricornio", "Tierro", "Onix Negro", "Loto", "Negro", "Saturno", "Plomo")
            PictureBoxgriego.Image = My.Resources.capricornio
        ElseIf (mes = 1 And dia >= 20) Or (mes = 2 And dia <= 18) Then
            setLabels("Acuario", "Aire", "Zafiro", "Tulipán", "Verde", "Urano", "Aluminio")
            PictureBoxgriego.Image = My.Resources.acuario
        ElseIf (mes = 2 And dia >= 19) Or (mes = 3 And dia <= 20) Then
            setLabels("Piscis", "Agua", "Amatista", "Violeta", "Violeta", "Neptuno", "Estaño")
            PictureBoxgriego.Image = My.Resources.pisis
        End If
    End Sub

    Private Sub setLabels(griego As String, elemento As String, piedra As String, flor As String, color As String, planeta As String, metal As String)
        lbl_griego.Text = griego
        lbl_elemento.Text = elemento
        lbl_piedra.Text = piedra
        lbl_flor.Text = flor
        lbl_planeta.Text = planeta
        lbl_metal.Text = metal
    End Sub

    Private Sub Btnobtener_Click(sender As Object, e As EventArgs) Handles Btnobtener.Click

    End Sub

    Private Sub obtenerSignoChino(anio As Integer)
        Dim resto As Integer
        resto = anio Mod 12
        If resto = 0 Then
            lbl_chino.Text = "Mono"
            PictureBoxchino.Image = My.Resources.mono
        ElseIf resto = 1 Then
            lbl_chino.Text = "Gallo"
            PictureBoxchino.Image = My.Resources.gallo
        ElseIf resto = 2 Then
            lbl_chino.Text = "Perro"
            PictureBoxchino.Image = My.Resources.perro
        ElseIf resto = 3 Then
            lbl_chino.Text = "Cerdo"
            PictureBoxchino.Image = My.Resources.cerdo
        ElseIf resto = 4 Then
            lbl_chino.Text = "Rata"
            PictureBoxchino.Image = My.Resources.rata
        ElseIf resto = 5 Then
            lbl_chino.Text = "Buey"
            PictureBoxchino.Image = My.Resources.buey
        ElseIf resto = 6 Then
            lbl_chino.Text = "Tigre"
            PictureBoxchino.Image = My.Resources.tigre
        ElseIf resto = 7 Then
            lbl_chino.Text = "Conejo"
            PictureBoxchino.Image = My.Resources.conejo
        ElseIf resto = 8 Then
            lbl_chino.Text = "Dragon"
            PictureBoxchino.Image = My.Resources.dragon
        ElseIf resto = 9 Then
            lbl_chino.Text = "Serpiente"
            PictureBoxchino.Image = My.Resources.serpiente
        ElseIf resto = 10 Then
            lbl_chino.Text = "Caballo"
            PictureBoxchino.Image = My.Resources.caballo
        ElseIf resto = 11 Then
            lbl_chino.Text = "Cabra"
            PictureBoxchino.Image = My.Resources.cabra
        End If
    End Sub

    Private Sub calcularNumeroAstral(dia As String, mes As String, anio As String)
        Dim numero As String
        Dim anioAux, suma As Integer

        If Val(anio) < 2000 Then
            anioAux = sumarNumeros(anio)
        Else
            anioAux = Val(Strings.Left(anio, 2)) + Val(Strings.Right(anio, 2))
        End If

        suma = Val(dia) + Val(mes) + Val(anioAux)
        Console.WriteLine(suma)
        While suma.ToString().Length <> 1
            suma = sumarNumeros(suma)
            Console.WriteLine(suma)
        End While


        numero = suma.ToString()
        lbl_numastral.Text = numero
    End Sub

    Private Function sumarNumeros(ByVal Cadena As String) As Integer
        Dim suma As Integer = 0
        For i = 1 To Len(Cadena)
            suma += Val(Strings.GetChar(Cadena, i))
        Next i
        Return suma
    End Function


    Private Function calcularEdad(minutoNacimiento As Integer, horaNacimiento As Integer, diaNacimiento As Integer, mesNacimiento As Integer, anioNacimiento As Integer, minActual As Integer, horaActual As Integer, diaActual As Integer, mesActual As Integer, anioActual As Integer)
        Dim minutosVividos As Integer
        minutosVividos = calcularMinutosVividos(minutoNacimiento, horaNacimiento, diaNacimiento, mesNacimiento, anioNacimiento, minActual, horaActual, diaActual, mesActual, anioActual)

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

    End Function

    Private Function calcularFechaGestacion(diaNac As Integer, mesNac As Integer, anioNac As Integer)
        Dim fecha As String
        Dim i, dia, mes, anio, maxDias As Integer
        fecha = ""

        dia = diaNac
        mes = mesNac
        anio = anioNac


        For i = 1 To 9
            mes -= 1
            If mes = 0 Then
                mes = 12
                anio -= 1
            End If
        Next i

        maxDias = obtenerDiasMes(mes, anio)

        If dia > maxDias Then
            dia = maxDias
        End If
        fecha = dia.ToString() + "/" + mes.ToString() + "/" + anio.ToString()
        lbl_gestacion.Text = fecha
    End Function

    Private Function calcularMinutosVividos(minutoNac As Integer, horaNac As Integer, diaNac As Integer, mesNac As Integer, anioNac As Integer, minActual As Integer, horaActual As Integer, diaActual As Integer, mesActual As Integer, anioActual As Integer) As Integer
        Dim maxDias As Integer
        Dim resultado As Integer
        resultado = 0


        While (minutoNac = minActual And horaActual = horaNac And diaNac = diaActual And mesActual = mesNac And anioActual = anioNac) = False

            If anioNac < anioActual Then

                If esAnioBiciesto(anioNac) Then
                    resultado += DiasAMinutos(366)
                Else
                    resultado += DiasAMinutos(365)
                End If
                anioNac += 1
            Else


                If Math.Abs(mesNac - mesActual) > 0 Then

                    If mesNac < mesActual Then
                        For i = mesNac To mesActual - 1
                            maxDias = obtenerDiasMes(i, anioNac)
                            resultado += DiasAMinutos(maxDias)
                        Next i
                    Else
                        For i = mesActual To mesNac - 1
                            maxDias = obtenerDiasMes(i, anioNac)
                            resultado -= DiasAMinutos(maxDias)
                        Next i
                    End If
                    mesNac = mesActual

                Else

                    If Math.Abs(diaActual - diaNac) > 0 Then
                        If diaActual < diaNac Then
                            resultado -= DiasAMinutos(Math.Abs(diaActual - diaNac))
                        Else
                            resultado += DiasAMinutos(Math.Abs(diaActual - diaNac))

                        End If
                        diaNac = diaActual
                    Else

                        If Math.Abs(horaNac - horaActual) > 0 Then

                            If horaNac < horaActual Then
                                resultado += Math.Abs((horaActual - horaNac)) * 60
                            Else
                                resultado -= Math.Abs((horaActual - horaNac)) * 60
                            End If
                            horaNac = horaActual
                        Else

                            If minutoNac <= minActual Then
                                resultado += Math.Abs(minActual - minutoNac)
                            Else
                                resultado -= Math.Abs(minActual - minutoNac)
                            End If

                            minutoNac = minActual
                        End If

                    End If


                End If


            End If




        End While




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
        For i = 1 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If "AEIOUÁÉÍÓÚÄËÏÖÜ".Contains(letra) Then
                Exit For
            End If
        Next i

        Return letra
    End Function
    Private Function ObtenerPrimerConsonante(ByVal Cadena As String) As String
        Dim letra As String = ""
        For i = 1 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If "AEIOUÁÉÍÓÚÄËÏÖÜ".Contains(letra) = False Then
                Exit For
            End If
        Next i
        Return letra
    End Function

    Private Function ObtenerPrimerVocalNoInicial(ByVal Cadena As String) As String
        Dim letra As String = ""
        For i = 2 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If "AEIOUÁÉÍÓÚÄËÏÖÜ".Contains(letra) Then
                Exit For
            End If
        Next i
        Return letra
    End Function
    Private Function ObtenerPrimerConsonanteNoInicial(ByVal Cadena As String) As String
        Dim letra As String = ""
        For i = 2 To Len(Cadena)
            letra = Mid(Cadena, i, 1)
            If "AEIOUÁÉÍÓÚÄËÏÖÜ".Contains(letra) = False Then
                Exit For
            End If
        Next i
        Return letra
    End Function

End Class


