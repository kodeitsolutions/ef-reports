﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grPagos_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class grPagos_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
            Dim lcParametro0HastaAux As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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
            loComandoSeleccionar.AppendLine("       CONVERT(NCHAR(10), Pagos.Fec_Ini, 112) As fec_ini,")
            loComandoSeleccionar.AppendLine("       CONVERT(NCHAR(10), Pagos.Fec_Ini, 103) As fec_ini2,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine(" 		    WHEN Detalles_Pagos.Tip_Ope = 'Efectivo' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Efectivo,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Ticket' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Ticket,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Cheque' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Cheque,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Tarjeta' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Deposito' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Depósito,")
            loComandoSeleccionar.AppendLine("       SUM(CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Pagos.Tip_Ope = 'Transferencia' THEN ISNULL(Detalles_Pagos.Mon_Net,0)")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END) AS Transferencia,")
            loComandoSeleccionar.AppendLine("       SUM(Detalles_Pagos.Mon_Net) As Total_Pagos")
			'loComandoSeleccionar.AppendLine(" INTO #temp")            
            loComandoSeleccionar.AppendLine(" FROM Pagos Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Pagos.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Pagos AS Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Pagos.Fec_ini BETWEEN DATEADD (DAY , -29, " & lcParametro0Hasta & " )")
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0HastaAux)
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
            
            loComandoSeleccionar.AppendLine(" GROUP BY CONVERT(NCHAR(10), Pagos.Fec_Ini, 112), CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)")
            loComandoSeleccionar.AppendLine(" ORDER BY CONVERT(NCHAR(10), Pagos.Fec_Ini, 112), CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)")
            loComandoSeleccionar.AppendLine(" ")
            
            
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grPagos_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrPagos_Fechas.ReportSource = loObjetoReporte

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
' CMS: 13/07/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
