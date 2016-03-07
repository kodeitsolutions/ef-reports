'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumen_SaldosCXP"
'-------------------------------------------------------------------------------------------'
Partial Class rResumen_SaldosCXP

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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' AND Cuentas_Pagar.Mon_Sal > 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) As Creditos_Pendientes,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal > 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) As Debitos_Pendientes,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal > 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) - SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' AND Cuentas_Pagar.Mon_Sal > 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) AS Saldo_Pendiente,")
            'loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal > 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) AS Saldo_Pendiente,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' AND Cuentas_Pagar.Mon_Sal = 0 Then Cuentas_Pagar.Mon_Net Else 0 End) As Creditos_Cancelados,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal = 0 Then Cuentas_Pagar.Mon_Net Else 0 End) As Debitos_Cancelados,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal = 0 Then Cuentas_Pagar.Mon_Net Else 0 End) - SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' AND Cuentas_Pagar.Mon_Sal = 0 Then Cuentas_Pagar.Mon_Net Else 0 End) AS Saldo_Cancelado,")
            'loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' AND Cuentas_Pagar.Mon_Sal = 0 Then Cuentas_Pagar.Mon_Sal Else 0 End) AS Saldo_Cancelado,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' Then Cuentas_Pagar.Mon_Net Else 0 End) As Creditos_Total,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' Then Cuentas_Pagar.Mon_Net Else 0 End) As Debitos_Total,")
            loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' Then Cuentas_Pagar.Mon_Net Else 0 End) - SUM(Case When Cuentas_Pagar.Tip_Doc = 'Credito' Then Cuentas_Pagar.Mon_Net Else 0 End) AS Saldo_Total")
            'loComandoSeleccionar.AppendLine(" 		 SUM(Case When Cuentas_Pagar.Tip_Doc = 'Debito' Then Cuentas_Pagar.Mon_Sal Else 0 End) AS Saldo_Total")
            loComandoSeleccionar.AppendLine(" FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores On Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini		Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip		Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro		Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven		Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Zon		Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Tip		Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Cla		Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tra		Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Mon		Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Suc		Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Rev		Between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("           And			" & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Cuentas_Pagar.Cod_Pro, Proveedores.Nom_Pro ")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
            
'me.mEscribirConsulta(loComandoSeleccionar.ToString)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumen_SaldosCXP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrResumen_SaldosCXP.ReportSource = loObjetoReporte

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
' CMS: 12/05/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'