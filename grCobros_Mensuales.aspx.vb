'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCobros_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class grCobros_Mensuales
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
            
            If cusAplicacion.goReportes.paParametrosIniciales(0) = 0 Then 
				lcParametro0Desde = "'" & Date.Now.Year & "'"
			End If
            
            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes," & lcParametro0Desde & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia INTO #temp_cobros")
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
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH, Cobros.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Cobros.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Efectivo,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Ticket,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Cheque,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Deposito,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0)) ELSE 0 END AS Transferencia")
            loComandoSeleccionar.AppendLine(" FROM Cobros Cobros")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       AND DATEPART(YEAR, Cobros.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            'loComandoSeleccionar.AppendLine("       AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Suc BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Cobros.fec_ini), DATEPART(MONTH, Cobros.fec_ini), Detalles_Cobros.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	#temp_cobros.Año AS Año,  ")
            loComandoSeleccionar.AppendLine("    	#temp_cobros.Mes AS Mes,  ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Deposito) AS Depósito, ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_cobros.Efectivo + #temp_cobros.Ticket + #temp_cobros.Cheque + #temp_cobros.Tarjeta + #temp_cobros.Deposito + #temp_cobros.Transferencia) AS Total_Cobros")
            loComandoSeleccionar.AppendLine(" FROM #temp_cobros  ")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_cobros.Año, #temp_cobros.Mes  ")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_cobros.Mes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_Mensuales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCobros_Mensuales.ReportSource = loObjetoReporte

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
' DLC: 30/04/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 03/09/2010: - Si por parametro el año es 0(cero), se toma el año en curso
'                  - Ajuste de la consulta para que tome solo los cobros distintos de Anulado
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'