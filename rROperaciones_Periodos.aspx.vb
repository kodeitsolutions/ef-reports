Imports System.Data
Imports cusAplicacion

Partial Class rROperaciones_Periodos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
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
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" EXEC sp_Resumen_Operaciones_por_Periodos")
            loComandoSeleccionar.AppendLine(" 	@sp_CodArt_Desde = " & lcParametro2Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodArt_Hasta = " & lcParametro2Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodCla_Desde = " & lcParametro4Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodCla_Hasta = " & lcParametro4Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodTip_Desde = " & lcParametro3Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodTip_Hasta = " & lcParametro3Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodCli_Desde = " & lcParametro0Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodCli_Hasta = " & lcParametro0Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodVen_Desde = " & lcParametro1Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodVen_Hasta = " & lcParametro1Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodDep_Desde = " & lcParametro5Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodDep_Hasta = " & lcParametro5Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_Status		 = '" & lcParametro6Desde & "',")
            loComandoSeleccionar.AppendLine(" 	@sp_TipRev		 = " & lcParametro7Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodRev_Desde = " & lcParametro8Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodRev_Hasta = " & lcParametro8Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodSuc_Desde = " & lcParametro9Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_CodSuc_Hasta = " & lcParametro9Hasta)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rROperaciones_Periodos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrROperaciones_Periodos.ReportSource = loObjetoReporte

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
' DLC: 27/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 21/07/2010: Optimización del procedimiento almacenado.
'-------------------------------------------------------------------------------------------'
