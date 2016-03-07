'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rECuentas_Clientes_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class rECuentas_Clientes_Ingles
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Hasta As String = cusAplicacion.goReportes.paParametrosFinales(9)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT           ")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli,     ")
            loComandoSeleccionar.AppendLine("		(SUM(CASE     ")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Cuentas_Cobrar.Mon_Sal      ")
            loComandoSeleccionar.AppendLine("				ELSE 0     ")
            loComandoSeleccionar.AppendLine("			END)      ")
            loComandoSeleccionar.AppendLine("		- SUM(CASE     ")
            loComandoSeleccionar.AppendLine("				WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal     ")
            loComandoSeleccionar.AppendLine("				ELSE 0     ")
            loComandoSeleccionar.AppendLine("		END)) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO	#tempSALDOINICIAL           ")
            loComandoSeleccionar.AppendLine("FROM	Clientes     ")
            loComandoSeleccionar.AppendLine("JOIN	Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli     ")
            loComandoSeleccionar.AppendLine("WHERE	 Cuentas_Cobrar.Fec_Ini < " & lcParametro1Desde & "  ")
            loComandoSeleccionar.AppendLine("        AND Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro0Hasta)
            'loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro1Desde)
            'loComandoSeleccionar.AppendLine("        AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Tra BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro7Hasta)

            If lcParametro9Hasta = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro8Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro8Desde)
            End If

            loComandoSeleccionar.AppendLine("        AND " & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine("GROUP BY Clientes.Cod_Cli    ")

            loComandoSeleccionar.AppendLine("SELECT    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Cli,     ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Fin,    ")
            loComandoSeleccionar.AppendLine("		CASE    ")
            loComandoSeleccionar.AppendLine("			WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' THEN Cuentas_Cobrar.Mon_Sal     ")
            loComandoSeleccionar.AppendLine("			ELSE 0    ")
            loComandoSeleccionar.AppendLine("		END AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("		CASE    ")
            loComandoSeleccionar.AppendLine("			WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal    ")
            loComandoSeleccionar.AppendLine("			ELSE 0    ")
            loComandoSeleccionar.AppendLine("		END AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Comentario    ")
            loComandoSeleccionar.AppendLine("INTO	#tempMOVIMIENTO    ")
            loComandoSeleccionar.AppendLine("FROM	Clientes    ")
            loComandoSeleccionar.AppendLine("JOIN	Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli    ")
            loComandoSeleccionar.AppendLine("WHERE	 Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("        AND Cuentas_Cobrar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Tra BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("     	 AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("        AND " & lcParametro7Hasta)

            If lcParametro9Hasta = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro8Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro8Desde)
            End If

            loComandoSeleccionar.AppendLine("        AND " & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine("SELECT  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Cod_Cli,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Nom_Cli,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Cod_Tip,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Documento,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Control,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Ini,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Fin,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Deb,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Hab,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Comentario,  	")
            loComandoSeleccionar.AppendLine("		ISNULL(#tempSALDOINICIAL.Sal_Ini,0) AS Sal_Ini,  	")
            loComandoSeleccionar.AppendLine("		0 AS Sal_Doc  	")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO  	")
            loComandoSeleccionar.AppendLine("LEFT JOIN	#tempSALDOINICIAL ON #tempSALDOINICIAL.Cod_Cli = #tempMOVIMIENTO.Cod_Cli  	")
            'loComandoSeleccionar.AppendLine("ORDER BY #tempMOVIMIENTO.Cod_Cli    ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & ", #tempMOVIMIENTO.Fec_Ini ASC ")


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmetne los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Cli", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Cli", GetType(String))
                loColumna.MaxLength = 200
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Documento", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Control", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Fin", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna.MaxLength = 500
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Cuenta_Actual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loNuevaFila.Item("Cod_Cli") = loFila("Cod_Cli")
                loNuevaFila.Item("Nom_Cli") = loFila("Nom_Cli")
                loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                loNuevaFila.Item("Documento") = loFila("Documento")
                loNuevaFila.Item("Control") = loFila("Control")
                loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                loNuevaFila.Item("Fec_Fin") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Fin")), "dd/MM/yyyy")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Comentario") = loFila("Comentario")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Cuenta_Actual = loFila("Cod_Cli")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Cli") <> Cuenta_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Cod_Cli") = loFila("Cod_Cli")
                    loNuevaFila.Item("Nom_Cli") = loFila("Nom_Cli")
                    loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                    loNuevaFila.Item("Documento") = loFila("Documento")
                    loNuevaFila.Item("Control") = loFila("Control")
                    loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                    loNuevaFila.Item("Fec_Fin") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Fin")), "dd/MM/yyyy")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Comentario") = loFila("Comentario")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Cuenta_Actual = loFila("Cod_Cli")

                    loTabla.AcceptChanges()

                Next lnNumeroFila


                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'

                '-------------------------------------------------------------------------------------------------------
                ' Verificando si el select (tabla nº0) trae registros
                '-------------------------------------------------------------------------------------------------------

                If (loDatosReporteFinal.Tables(0).Rows.Count <= 0) Then
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                              "No se Encontraron Registros para los Parámetros Especificados. ", _
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                               "350px", _
                                               "200px")
                End If


                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rECuentas_Clientes_Ingles", loDatosReporteFinal)
            Else

                ''-------------------------------------------------------------------------------------------------------
                '' Verificando si el select (tabla nº0) trae registros
                ''-------------------------------------------------------------------------------------------------------

                If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                              "No se Encontraron Registros para los Parámetros Especificados. ", _
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                               "350px", _
                                               "200px")
                End If

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rECuentas_Clientes_Ingles", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrECuentas_Clientes_Ingles.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try


    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub
End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' Douglas Cortez: 12/05/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
