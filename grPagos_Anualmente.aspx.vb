'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grPagos_Anualmente"
'-------------------------------------------------------------------------------------------'
Partial Class grPagos_Anualmente
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
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT")
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
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Depósito' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Depósito,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Transferencia")
            loComandoSeleccionar.AppendLine(" INTO #temp_Pagos")
            loComandoSeleccionar.AppendLine(" FROM Pagos Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Pagos.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Pagos AS Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Pagos.Fec_ini BETWEEN DATEADD (YYYY , -10, " & lcParametro0Hasta & " )")
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
              
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Rev BETWEEN " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Rev NOT BETWEEN " & lcParametro6Desde)
            End If
            
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Pagos.fec_ini), Detalles_Pagos.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ORDER BY DATEPART(YEAR, Pagos.fec_ini) ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	Año,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Depósito) AS Depósito, ")
            loComandoSeleccionar.AppendLine("    	SUM(Transferencia) AS Transferencia ")
            loComandoSeleccionar.AppendLine(" INTO #temp_totPagos  ")
            loComandoSeleccionar.AppendLine(" FROM #temp_Pagos  ")
            loComandoSeleccionar.AppendLine(" GROUP BY Año")
            loComandoSeleccionar.AppendLine(" ORDER BY Año ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("       #temp_totPagos.Año AS Año,    ")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Efectivo) AS Efectivo,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Ticket) AS Ticket,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Cheque)   AS Cheque,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Tarjeta)  AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Depósito) AS Depósito,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totPagos.Efectivo + #temp_totPagos.Cheque + #temp_totPagos.Tarjeta + #temp_totPagos.Depósito + #temp_totPagos.Transferencia + #temp_totPagos.Ticket) AS Total_Pagos")
            loComandoSeleccionar.AppendLine(" Into #Final   ")
            loComandoSeleccionar.AppendLine(" FROM	#temp_totPagos  ")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_totPagos.Año")
            
            loComandoSeleccionar.AppendLine(" Union All ")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -10, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -9, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -8, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -7, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -6, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -5, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -4, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -3, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -2, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -1, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, " & lcParametro0Hasta & ")  As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Depósito, 0 As Transferencia, 0 As Total_Pagos")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_totPagos.Año ")
            
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		Año,     ")
            loComandoSeleccionar.AppendLine(" 		SUM(Efectivo) AS Efectivo, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Ticket) AS Ticket, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Cheque)   AS Cheque, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Tarjeta)  AS Tarjeta, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Depósito) AS Depósito, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Transferencia) AS Transferencia, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Total_Pagos) AS Total_Pagos ")
            loComandoSeleccionar.AppendLine(" FROM #Final ")
			loComandoSeleccionar.AppendLine(" WHERE Año BETWEEN")
			loComandoSeleccionar.AppendLine(" (Case ")
			loComandoSeleccionar.AppendLine(" 	When  " & lcParametro0Desde & " > DATEADD(YYYY , -10, " & lcParametro0Hasta & ") Then Datepart(yyyy,DATEADD(YYYY , -10, " & lcParametro0Hasta & "))")
			loComandoSeleccionar.AppendLine(" 	Else Datepart(yyyy," & lcParametro0Desde & ")")
			loComandoSeleccionar.AppendLine(" 	End )")
			loComandoSeleccionar.AppendLine("   AND Datepart(yyyy,'20100611 23:59:59.998')")
            
            loComandoSeleccionar.AppendLine(" GROUP BY Año ")
            loComandoSeleccionar.AppendLine(" ORDER BY Año  ")

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
            
            Dim Total As Decimal = 0
            
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                Total = Total + loFilas.Item("Total_Pagos")

            Next loFilas
            
            If Total = 0 And laDatosReporte.Tables(0).Rows.Count >= 0 Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If
            

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grPagos_Anualmente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrPagos_Anualmente.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.message, _
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
' CMS: 25/06/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
