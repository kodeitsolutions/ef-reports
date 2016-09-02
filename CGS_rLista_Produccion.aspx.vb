'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLista_Produccion"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLista_Produccion
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcFechaIni		AS DATETIME		= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcFechaFin		AS DATETIME		= " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProduccionIni	AS VARCHAR(10)	= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProduccionFin	AS VARCHAR(10)	= " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcArtIni			AS VARCHAR(8)	= " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcArtFin			AS VARCHAR(8)	= " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteIni			AS VARCHAR(30)	= " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcLoteFin			AS VARCHAR(30)	= " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Encabezados.Proyecto					AS Produccion,")
            loComandoSeleccionar.AppendLine("		Renglones.Cod_Art						AS Articulo,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Articulos.Atributo_A = ' ' ")
            loComandoSeleccionar.AppendLine("			 THEN CAST(0 AS DECIMAL(28,2)) * Renglones.Can_Art")
            loComandoSeleccionar.AppendLine("			 ELSE CONVERT(NUMERIC(18,2), REPLACE(Articulos.Atributo_A,',','.') )  * Renglones.Can_Art")
            loComandoSeleccionar.AppendLine("		END					                    AS Desp_Standard")
            loComandoSeleccionar.AppendLine("INTO #tmpConsumo")
            loComandoSeleccionar.AppendLine("FROM Encabezados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones ON Encabezados.Documento = Renglones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Encabezados.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("   AND Encabezados.Fec_Ini BETWEEN @lcFechaIni AND @lcFechaFin")
            loComandoSeleccionar.AppendLine("	AND Encabezados.Proyecto BETWEEN @lcProduccionIni AND @lcProduccionFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Proyectos.Cod_Pro									AS Produccion,")
            loComandoSeleccionar.AppendLine("		Formulas.Cod_Art									AS Cod_Articulo,")
            loComandoSeleccionar.AppendLine("		Formulas.Nom_Art									AS Articulo,")
            loComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cod_Lot, '')				AS Lote,")
            loComandoSeleccionar.AppendLine("		Reng_OTrabajo.Can_Art								AS Obtenido,")
            loComandoSeleccionar.AppendLine("		SUM(Reng_Consumo.Can_Art)							AS Consumido,")
            loComandoSeleccionar.AppendLine("		SUM(Reng_Consumo.Can_Art) - Reng_OTrabajo.Can_Art	AS Desp_Obtenido,")
            loComandoSeleccionar.AppendLine("		SUM(#tmpConsumo.Desp_Standard)						AS Desp_Standard,")
            loComandoSeleccionar.AppendLine("		((SUM(Reng_Consumo.Can_Art) - Reng_OTrabajo.Can_Art) / SUM(Reng_Consumo.Can_Art)) * 100 AS PorcDesp_Obt,")
            loComandoSeleccionar.AppendLine("		(SUM(#tmpConsumo.Desp_Standard) / SUM(Reng_Consumo.Can_Art)) * 100						AS PorcDesp_Std")
            loComandoSeleccionar.AppendLine("FROM Proyectos")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Proyectos ON Proyectos.Cod_Pro = Renglones_Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Formulas ON Formulas.Documento = Renglones_Proyectos.Cod_Par")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Orden_Trabajo ON Orden_Trabajo.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Orden_Trabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Reng_OTrabajo ON Reng_OTrabajo.Documento = Orden_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Reng_OTrabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("		AND Reng_OTrabajo.Cod_Reg = Formulas.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Orden_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Formulas.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Consumo ON Consumo.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Consumo.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Reng_Consumo ON Consumo.Documento = Reng_Consumo.Documento")
            loComandoSeleccionar.AppendLine("		AND Reng_Consumo.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN #tmpConsumo ON Proyectos.Cod_Pro = #tmpConsumo.Produccion")
            loComandoSeleccionar.AppendLine("		AND #tmpConsumo.Articulo = Reng_Consumo.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Proyectos.Fec_Ini BETWEEN @lcFechaIni AND @lcFechaFin")
            loComandoSeleccionar.AppendLine("	AND Proyectos.Cod_Pro BETWEEN @lcProduccionIni AND @lcProduccionFin")
            loComandoSeleccionar.AppendLine("	AND Formulas.Cod_Art BETWEEN @lcArtIni AND @lcArtFin")
            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("	AND Operaciones_Lotes.Cod_Lot BETWEEN @lcLoteIni AND @lcLoteFin")
            End If
            loComandoSeleccionar.AppendLine("GROUP BY Proyectos.Cod_Pro, Formulas.Cod_Art, Formulas.Nom_Art, Operaciones_Lotes.Cod_Lot, Reng_OTrabajo.Can_Art")
            loComandoSeleccionar.AppendLine("ORDER BY Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpConsumo")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLista_Produccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rLista_Produccion.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

