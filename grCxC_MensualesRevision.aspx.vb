'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System.Data.SqlClient

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCxC_MensualesRevision"
'-------------------------------------------------------------------------------------------'
Partial Class grCxC_MensualesRevision
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" EXEC sp_CalculoCxC_MensualRevision")
            loComandoSeleccionar.AppendLine(" 		@sp_lcFecIni = " & lcParametro0Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcFecFin = " & lcParametro0Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxC_CodTip_Desde = " & lcParametro8Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxC_CodTip_Hasta = " & lcParametro8Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxC_Status = " & lcParametro9Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodCli_Desde = " & lcParametro2Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodCli_Hasta = " & lcParametro2Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodVen_Desde = " & lcParametro1Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodVen_Hasta = " & lcParametro1Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodMon_Desde = " & lcParametro7Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodMon_Hasta = " & lcParametro7Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_TipRev = '" & lcParametro11Desde & "',")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodRev_Desde = " & lcParametro10Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodRev_Hasta = " & lcParametro10Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodSuc_Desde = " & lcParametro6Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcCxCC_CodSuc_Hasta = " & lcParametro6Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodTip_Desde = " & lcParametro4Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodTip_Hasta = " & lcParametro4Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodZon_Desde = " & lcParametro5Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodZon_Hasta = " & lcParametro5Hasta & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodCla_Desde = " & lcParametro3Desde & ",")
            loComandoSeleccionar.AppendLine(" 		@sp_lcClientes_CodCla_Hasta = " & lcParametro3Hasta)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCxC_MensualesRevision", laDatosReporte)

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If laDatosReporte.Tables(0).Rows.Count <= 0 Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If



            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCxC_MensualesRevision.ReportSource = loObjetoReporte

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
' DLC: 07/09/2010: Codigo inicial (tomando como base el reporte grCxC_MensualesSucursal)
'-------------------------------------------------------------------------------------------'