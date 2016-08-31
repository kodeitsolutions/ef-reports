'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLProduccion_Desperdicio"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLProduccion_Desperdicio

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcArticuloIni AS VARCHAR(8) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArticuloFin AS VARCHAR(8) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteIni	    AS VARCHAR(30) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteFin	    AS VARCHAR(30) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Formulas.Cod_Art			AS CodArt_Elaborado,")
            loComandoSeleccionar.AppendLine("		Formulas.Nom_Art			AS NomArt_Elaborado,")
            loComandoSeleccionar.AppendLine("		Formulas.Cod_Uni			AS Unidad,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot	AS Lote,")
            loComandoSeleccionar.AppendLine("		Proyectos.Cod_Pro			AS Orden_Produccion,")
            loComandoSeleccionar.AppendLine("		Renglones.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Renglones.Nom_Art			AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones.Can_Art			AS Cantidad,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Articulos.Atributo_A = ' ' ")
            loComandoSeleccionar.AppendLine("			 THEN CAST(0 AS DECIMAL(28,2)) * Renglones.Can_Art")
            loComandoSeleccionar.AppendLine("			 ELSE CONVERT(NUMERIC(18,2), REPLACE(Articulos.Atributo_A,',','.') )  * Renglones.Can_Art")
            loComandoSeleccionar.AppendLine("		END							AS Desperdicio_Stand,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Articulos.Atributo_A = ' ' ")
            loComandoSeleccionar.AppendLine("			 THEN CAST(0 AS DECIMAL(28,2)) ")
            loComandoSeleccionar.AppendLine("			 ELSE CONVERT(NUMERIC(18,2), REPLACE(Articulos.Atributo_A,',','.') ) * 100 ")
            loComandoSeleccionar.AppendLine("		END							AS Desperdicio,")
            loComandoSeleccionar.AppendLine("		Renglones_OTrabajo.Can_Art	AS Obtenido,")
            loComandoSeleccionar.AppendLine("		(SELECT SUM(Consumo_P.Can_Art) FROM Renglones AS Consumo_P")
            loComandoSeleccionar.AppendLine("		WHERE Consumo_P.Documento = Consumo.Documento AND Consumo_P.Origen = 'Consumos Produccion')		AS Consumido")
            loComandoSeleccionar.AppendLine("FROM Renglones")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Consumo ON Consumo.Documento = Renglones.Documento")
            loComandoSeleccionar.AppendLine("		AND Consumo.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Proyectos ON Consumo.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Orden_Trabajo ON Proyectos.Cod_Pro = Orden_Trabajo.Proyecto")
            loComandoSeleccionar.AppendLine("		AND Orden_Trabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_OTrabajo ON Renglones_OTrabajo.Documento = Orden_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OTrabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Formulas ON Renglones_OTrabajo.Cod_Reg = Formulas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Orden_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Formulas.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Formulas.Cod_Art BETWEEN @lcArticuloIni AND @lcArticuloFin")
            loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot BETWEEN @lcLoteIni AND @lcLoteFin")

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

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLProduccion_Desperdicio", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rLProduccion_Desperdicio.ReportSource = loObjetoReporte

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
' JJD: 09/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'