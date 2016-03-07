'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSMensual_Conciliado"
'-------------------------------------------------------------------------------------------'
Partial Class rSMensual_Conciliado

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lnParametro2Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(2)
            Dim lnParametro2Hasta As Integer = cusAplicacion.goReportes.paParametrosFinales(2)
            Dim lnParametro3Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(3)

            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(lnParametro2Desde)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(lnParametro2Hasta)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(lnParametro3Desde)

            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))

            If lnParametro2Desde = 0 Then
                lnParametro2Desde = Date.Now().Month()
                lcParametro2Desde = goServicios.mObtenerCampoFormatoSQL(lnParametro2Desde)
                lcParametro2Hasta = goServicios.mObtenerCampoFormatoSQL(lnParametro2Desde)
            End If

            If lnParametro3Desde = 0 Then
                lnParametro3Desde = Date.Now().Year()
                lcParametro3Desde = goServicios.mObtenerCampoFormatoSQL(lnParametro3Desde)
            End If

            Dim lcParametro0Fecha As String = goServicios.mObtenerCampoFormatoSQL(New Date(lnParametro3Desde, lnParametro2Desde, 1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Bancarias.Cod_Cue, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine(" 			MONTH(Movimientos_Cuentas.Fec_Cil)                                  As  Mes, ")
            loComandoSeleccionar.AppendLine(" 			ISNULL((Mon_Hab - Mon_Deb - Mon_Imp1), CAST(0 AS Decimal (28,10)))  As  Sal_Con, ")
            loComandoSeleccionar.AppendLine(" 			CAST(0 AS Decimal (28,10))                                          As  Sal_Act, ")
            loComandoSeleccionar.AppendLine(" 			" & lcParametro3Desde & "                                           As  Año ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpSaldosCuentas01 ")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Bancarias LEFT JOIN Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" 			ON Cuentas_Bancarias.Cod_Cue            =   Movimientos_Cuentas.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" WHERE		Cuentas_Bancarias.Cod_Cue               BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Con         BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And MONTH(Movimientos_Cuentas.Fec_Cil)  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Movimientos_Cuentas.Fec_Cil)   =   " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status          IN  (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Doc_Cil         =	'1' ")
            loComandoSeleccionar.AppendLine(" ORDER BY	1, 2 ")

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Bancarias.Cod_Cue, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine(" 			MONTH(Movimientos_Cuentas.Fec_Ini)                                  As  Mes, ")
            loComandoSeleccionar.AppendLine(" 			CAST(0 AS Decimal (28,10))                                          As  Sal_Con, ")
            loComandoSeleccionar.AppendLine(" 			ISNULL((Mon_Hab - Mon_Deb - Mon_Imp1), CAST(0 AS Decimal (28,10)))  AS  Sal_Act, ")
            loComandoSeleccionar.AppendLine(" 			" & lcParametro3Desde & "                                           As  Año ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpSaldosCuentas02 ")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Bancarias LEFT JOIN Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" 			ON Cuentas_Bancarias.Cod_Cue            =   Movimientos_Cuentas.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" WHERE		Cuentas_Bancarias.Cod_Cue               BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Con         BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And MONTH(Movimientos_Cuentas.Fec_Ini)  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Movimientos_Cuentas.Fec_Ini)   =   " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status          IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY	1, 2 ")

            loComandoSeleccionar.AppendLine(" SELECT    * ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpSaldosCuentasFinal ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpSaldosCuentas01 ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT    * ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpSaldosCuentas02 ")

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           Num_Cue, ")
            loComandoSeleccionar.AppendLine(" 			Año, ")
            loComandoSeleccionar.AppendLine(" 			Mes, ")

            'loComandoSeleccionar.AppendLine(" 			CASE WHEN ")
            'loComandoSeleccionar.AppendLine(" 		    	        Mes = '1' THEN '01'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '2' THEN '02'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '3' THEN '03'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '4' THEN '04'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '5' THEN '05'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '6' THEN '06'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '7' THEN '07'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '8' THEN '08'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '9' THEN '09'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '10' THEN '10'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '11' THEN '11'")
            'loComandoSeleccionar.AppendLine(" 		    	ELSE    Mes = '12' THEN '12'")
            'loComandoSeleccionar.AppendLine(" 			END             As  Nom_Mes, ")

            loComandoSeleccionar.AppendLine(" 			Sum(Sal_Con)	As  Sal_Con, ")
            loComandoSeleccionar.AppendLine(" 			Sum(Sal_Act)	As  Sal_Act ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpSaldosCuentasFinal ")
            loComandoSeleccionar.AppendLine(" GROUP BY	Cod_Cue, Nom_Cue, Num_Cue, Año, Mes ")
            loComandoSeleccionar.AppendLine(" ORDER BY	Cod_Cue, Nom_Cue, Num_Cue, Año, Mes ")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSMensual_Conciliado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSMensual_Conciliado.ReportSource = loObjetoReporte

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
' JJD: 06/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
