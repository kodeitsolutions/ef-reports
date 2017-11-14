Imports System.Data
Partial Class KDE_rLibro_Diario
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Comprobantes.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cue			AS Cod_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Nom_Cue			AS Nom_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux			AS Auxiliar, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Deb			AS Debe, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Hab			AS Haber, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Doc_Ori			AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Tip_Ori			AS Tip_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Comentario		AS Comentario,")
            loComandoSeleccionar.AppendLine("       CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha")
            loComandoSeleccionar.AppendLine("FROM Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Comprobantes ON Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("WHERE Comprobantes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("ORDER BY Comprobantes.Documento, Renglones_Comprobantes.Renglon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("KDE_rLibro_Diario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvKDE_rLibro_Diario.ReportSource = loObjetoReporte

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
' KDE: 14/11/17: Codigo inicial
'-------------------------------------------------------------------------------------------'