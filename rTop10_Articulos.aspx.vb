'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTop10_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rTop10_Articulos
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("EXEC sp_Calculo_Top10_Articulos")
            loComandoSeleccionar.AppendLine("       @sp_FecIni          = " & lcParametro0Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_FecFin          = " & lcParametro0Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodArt_Desde    = " & lcParametro1Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodArt_Hasta    = " & lcParametro1Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodTip_Desde    = " & lcParametro3Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodTip_Hasta    = " & lcParametro3Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCla_Desde    = " & lcParametro2Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodCla_Hasta    = " & lcParametro2Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_Status          = " & lcParametro13Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodAlm_Desde    = " & lcParametro4Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodAlm_Hasta    = " & lcParametro4Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodDep_Desde    = " & lcParametro5Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodDep_Hasta    = " & lcParametro5Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSec_Desde    = " & lcParametro6Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSec_Hasta    = " & lcParametro6Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodMar_Desde    = " & lcParametro7Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodMar_Hasta    = " & lcParametro7Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodPro_Desde    = " & lcParametro8Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodPro_Hasta    = " & lcParametro8Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodUni_Desde    = " & lcParametro9Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodUni_Hasta    = " & lcParametro9Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSuc_Desde    = " & lcParametro10Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodSuc_Hasta    = " & lcParametro10Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_TipRev          = " & lcParametro11Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodRev_Desde    = " & lcParametro12Desde & ",")
            loComandoSeleccionar.AppendLine("       @sp_CodRev_Hasta    = " & lcParametro12Hasta & ",")
            loComandoSeleccionar.AppendLine("       @sp_NumDecCant      = " & goOpciones.pnDecimalesParaCantidad())

            'Response.Clear()
            'Response.Write("<html><body><pre>" & vbNewLine)
            'Response.Write(loComandoSeleccionar.ToString)
            'Response.Write("</pre></body></html>")
            'Response.Flush()
            'Response.End()

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTop10_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTop10_Articulos.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & ", StackTrace: " & loExcepcion.StackTrace, _
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
' DLC: 01/07/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 06/07/2010: La consulta a la base de datos, se pasó a ejecutar a traves de 
'                    un procedimiento almacenado
'-------------------------------------------------------------------------------------------'