Imports System.Data
Partial Class CGS_rComprobante_Diario
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcDoc_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcDoc_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAux_Desde AS VARCHAR(12) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAux_Hasta AS VARCHAR(12) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Comprobantes.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("		Comprobantes.Fec_Ini					AS Fecha_Comprobante,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon			AS Renglon,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cue			AS Cod_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Nom_Cue			AS Nom_Cuenta, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux			AS Auxiliar, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Deb			AS Debe, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Hab			AS Haber, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Doc_Ori			AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Tip_Ori			AS Tip_Ori, ")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Comentario		AS Comentario,")
            loComandoSeleccionar.AppendLine("       CASE WHEN @lcDoc_Desde <> '' ")
            loComandoSeleccionar.AppendLine("	         THEN CONCAT(@lcDoc_Desde, ' - ', @lcDoc_Hasta)")
            loComandoSeleccionar.AppendLine("	         ELSE 'N/E' END						AS Docs,")
            loComandoSeleccionar.AppendLine("       CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("       CASE WHEN @lcCodAux_Desde <> ''")
            loComandoSeleccionar.AppendLine("		     THEN (SELECT Nom_Aux FROM Auxiliares WHERE Cod_Aux = @lcCodAux_Desde)")
            loComandoSeleccionar.AppendLine("		     ELSE '' END				        AS Aux_Desde,")
            loComandoSeleccionar.AppendLine("       CASE WHEN @lcCodAux_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		     THEN (SELECT Nom_Aux FROM Auxiliares WHERE Cod_Aux = @lcCodAux_Hasta)")
            loComandoSeleccionar.AppendLine("		     ELSE '' END				        AS Aux_Hasta")
            loComandoSeleccionar.AppendLine("FROM Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Comprobantes ON Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("WHERE Comprobantes.Documento BETWEEN @lcDoc_Desde AND @lcDoc_Hasta")
            loComandoSeleccionar.AppendLine("	AND Comprobantes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Doc_Ori IN (SELECT RC.Doc_Ori ")
            loComandoSeleccionar.AppendLine("										  FROM Renglones_Comprobantes AS RC")
            loComandoSeleccionar.AppendLine("										  WHERE RC.Cod_Aux BETWEEN @lcCodAux_Desde AND @lcCodAux_Hasta")
            loComandoSeleccionar.AppendLine("											AND RC.Tip_Ori = Renglones_Comprobantes.Tip_Ori)")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rComprobante_Diario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rComprobante_Diario.ReportSource = loObjetoReporte

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
' JJD: 23/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  17/08/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT:  16/05/11: Mejora de la vista de Diseño
'-------------------------------------------------------------------------------------------'