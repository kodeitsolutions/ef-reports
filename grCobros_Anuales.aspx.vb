'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCobros_Anuales"
'-------------------------------------------------------------------------------------------'
Partial Class grCobros_Anuales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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

            Dim lcDateDesde As Date = New System.DateTime(Val(Mid(lcParametro0Desde, 2, 4)), Val(Mid(lcParametro0Desde, 6, 2)), Val(Mid(lcParametro0Desde, 8, 2)), 0, 0, 0)
            Dim lcDateHasta As Date = New System.DateTime(Val(Mid(lcParametro0Hasta, 2, 4)), Val(Mid(lcParametro0Hasta, 6, 2)), Val(Mid(lcParametro0Hasta, 8, 2)), 23, 59, 59)
            Dim lcNumAños As Integer = (Year(lcDateHasta) - Year(lcDateDesde)) + 1

            Dim lcAño As Integer = Year(lcDateDesde)

            If lcNumAños > 30 Then
                lcAño = Year(lcDateHasta) - 29
                lcDateDesde = New System.DateTime(lcAño, 1, 1, 0, 0, 0)
                lcNumAños = (Year(lcDateHasta) - lcAño) + 1
            End If

            loComandoSeleccionar.AppendLine(" SELECT " & lcAño & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia INTO #temp_cobros")
            For lcIndex As Integer = 1 To lcNumAños - 1
                lcAño = lcAño + 1
                loComandoSeleccionar.AppendLine(" UNION ALL")
                loComandoSeleccionar.AppendLine(" SELECT " & lcAño & " AS Año, 0 As Efectivo, 0 AS Ticket, 0 AS Cheque, 0 AS Tarjeta, 0 AS Deposito, 0 AS Transferencia")
            Next lcIndex

            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Cobros.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine(" 		    WHEN Detalles_Cobros.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Efectivo,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Ticket,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Cheque,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Deposito,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Transferencia")
            loComandoSeleccionar.AppendLine(" FROM Cobros Cobros")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" WHERE DATEPART(YEAR,Cobros.Fec_ini) BETWEEN " & Year(lcDateDesde))
            loComandoSeleccionar.AppendLine("       AND " & Year(lcDateHasta))
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Cobros.fec_ini), Detalles_Cobros.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	#temp_cobros.Año AS Año,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Deposito) AS Depósito, ")
            loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_cobros.Efectivo + #temp_cobros.Ticket + #temp_cobros.Cheque + #temp_cobros.Tarjeta + #temp_cobros.Deposito + #temp_cobros.Transferencia) AS Total_Cobros")
            loComandoSeleccionar.AppendLine(" FROM #temp_cobros  ")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_cobros.Año")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_cobros.Año")

            'Response.Clear()
            'Response.Write("<html><body><pre>" & vbNewLine)
            'Response.Write(loComandoSeleccionar.ToString)
            'Response.Write("</pre></body></html>")
            'Response.Flush()
            'Response.End()

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_Anuales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCobros_Anuales.ReportSource = loObjetoReporte

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
' Douglas Cortez: 17/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'