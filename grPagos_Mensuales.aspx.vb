'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grPagos_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class grPagos_Mensuales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia INTO #tempPAGOS")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH, Pagos.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Pagos.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine(" 		    WHEN Detalles_Pagos.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Efectivo,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Ticket,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Cheque,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Deposito,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Transferencia")
            loComandoSeleccionar.AppendLine(" FROM Pagos Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Pagos.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Pagos AS Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE DATEPART(YEAR, Pagos.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Pagos.fec_ini), DATEPART(MONTH, Pagos.fec_ini), Detalles_Pagos.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	#tempPAGOS.Año AS Año,  ")
            loComandoSeleccionar.AppendLine("    	#tempPAGOS.Mes AS Mes,  ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempPAGOS.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Deposito) AS Depósito, ")
            loComandoSeleccionar.AppendLine("    	SUM(#tempPAGOS.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine(" 		SUM(#tempPAGOS.Efectivo + #tempPAGOS.Ticket + #tempPAGOS.Cheque + #tempPAGOS.Tarjeta + #tempPAGOS.Deposito + #tempPAGOS.Transferencia) AS Total_Pagos")
            loComandoSeleccionar.AppendLine(" FROM #tempPAGOS  ")
            loComandoSeleccionar.AppendLine(" GROUP BY #tempPAGOS.Año, #tempPAGOS.Mes  ")
            loComandoSeleccionar.AppendLine(" ORDER BY #tempPAGOS.Mes")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grPagos_Mensuales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrPagos_Mensuales.ReportSource = loObjetoReporte

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
' Douglas Cortez 12/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
