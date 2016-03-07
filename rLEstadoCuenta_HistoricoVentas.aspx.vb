﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLEstadoCuenta_HistoricoVentas"
'-------------------------------------------------------------------------------------------'
Partial Class rLEstadoCuenta_HistoricoVentas
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
            'Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("EXEC   sp_Estado_de_Cuenta_Ventas")
            'If (lcParametro5Desde = "Si") Then
            '    ' El reporte SI lleva detalle de cobros
            '    loComandoSeleccionar.AppendLine("       @sp_LisDetallado    = 1,")
            'Else
            '    ' El reporte NO lleva detalle de cobros
            '    loComandoSeleccionar.AppendLine("       @sp_LisDetallado    = 0,")
            'End If
            loComandoSeleccionar.AppendLine("       @sp_FecIni          = " & lcParametro0Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_FecFin          = " & lcParametro0Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCli_Desde    = " & lcParametro1Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCli_Hasta    = " & lcParametro1Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCla_Desde    = " & lcParametro2Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCla_Hasta    = " & lcParametro2Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodTip_Desde    = " & lcParametro3Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodTip_Hasta    = " & lcParametro3Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodZon_Desde    = " & lcParametro5Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodZon_Hasta    = " & lcParametro5Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodVen_Desde    = " & lcParametro4Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodVen_Hasta    = " & lcParametro4Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodMon_Desde    = " & lcParametro6Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodMon_Hasta    = " & lcParametro6Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSuc_Desde    = " & lcParametro7Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSuc_Hasta    = " & lcParametro7Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_TipRev          = " & lcParametro8Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodRev_Desde    = " & lcParametro9Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodRev_Hasta    = " & lcParametro9Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_Ordenamiento    = '" & lcOrdenamiento & "'")

           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLEstadoCuenta_HistoricoVentas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLEstadoCuenta_HistoricoVentas.ReportSource = loObjetoReporte

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
' DLC: 21/07/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' DLC: 02/09/2010: Cambio de la consulta a procedimiento almacenado.
'                   - Se ajusto los filtro de las consultas.
'                   - Se ajusto la forma en que se calcula los cobros.
'                   - Se ajusto la forma en que se calcula las cuentas por cobrar.
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Cobros, asi como también,
'                ajustar en el RPT, la forma de mostrar los detalles de Cobros.
'-------------------------------------------------------------------------------------------'
' MAT: 13/04/11: Reprogramación del Reporte y su respectivo Store Procedure
'-------------------------------------------------------------------------------------------'
' MAT: 14/04/11: Ajuste de la vista de Diseño.
'-------------------------------------------------------------------------------------------'
' MAT: 27/04/11: Se elimino el filtro Detalle
'-------------------------------------------------------------------------------------------'

