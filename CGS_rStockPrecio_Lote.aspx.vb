'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rStockPrecio_Lote"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rStockPrecio_Lote
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcLote AS VARCHAR(30) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Lotes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Cod_uni1, ")
            loComandoSeleccionar.AppendLine("       Lotes.Cod_Lot, ")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Pro1,")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Pro2,")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Ult1,")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Ult2,")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Ant1,")
            loComandoSeleccionar.AppendLine("		Lotes.Cos_Ant2,")
            loComandoSeleccionar.AppendLine("       Lotes.Exi_Act1 ")
            loComandoSeleccionar.AppendLine("FROM Lotes ")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Lotes.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Usa_Lot = 1 ")
            loComandoSeleccionar.AppendLine("	AND Lotes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("	AND Lotes.Cod_Lot LIKE '%'+@lcLote+'%'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("ORDER BY  Lotes.Cod_Art, Lotes.Cod_Lot")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rStockPrecio_Lote", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rStockPrecio_Lote.ReportSource = loObjetoReporte

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
' CMS: 28/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'