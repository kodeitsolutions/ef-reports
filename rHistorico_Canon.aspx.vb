'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rHistorico_Canon"
'-------------------------------------------------------------------------------------------'
Partial Class rHistorico_Canon
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Auditorias.Cod_Usu, Auditorias.Registro, Auditorias.Codigo, Auditorias.Accion,")
            loComandoSeleccionar.AppendLine("			CAST(LOWER(CAST(Auditorias.Detalle AS NVARCHAR(MAX))) AS XML) AS Detalle")
            loComandoSeleccionar.AppendLine("INTO		#tmpCambios")
            loComandoSeleccionar.AppendLine("FROM		Auditorias")
            loComandoSeleccionar.AppendLine("WHERE		Auditorias.Tabla = 'Articulos'")
            loComandoSeleccionar.AppendLine("		AND Auditorias.Cod_Obj <> 'Piraña'")
            loComandoSeleccionar.AppendLine("		AND Auditorias.Tipo = 'Datos'")
            loComandoSeleccionar.AppendLine("	    AND Auditorias.Codigo BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("	    AND Auditorias.Registro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	Auditorias.Codigo, Auditorias.Registro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tmpCambios.Cod_Usu, #tmpCambios.Registro,  ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli			AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("		Clientes.Rif				AS Rif_Cli, ")
            loComandoSeleccionar.AppendLine("		#tmpCambios.Codigo			AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("		#tmpCambios.Accion, ")
            loComandoSeleccionar.AppendLine("		Detalle.value('(/detalle/campos/campo[@nombre=""precio1""]/antes)[1]', 'VARCHAR(20)')		AS Canon_Anterior,")
            loComandoSeleccionar.AppendLine("		Detalle.value('(/detalle/campos/campo[@nombre=""precio1""]/despues)[1]', 'VARCHAR(20)')		AS Canon_Nuevo,")
            loComandoSeleccionar.AppendLine("		Detalle.value('(/detalle/campos/campo[@nombre=""fec_pre1""]/despues)[1]', 'VARCHAR(10)')	AS Fecha_Inicio,")
            loComandoSeleccionar.AppendLine("		Detalle.value('(/detalle/campos/campo[@nombre=""fec_pre2""]/despues)[1]', 'VARCHAR(10)')	AS Fecha_Vencimiento")
            loComandoSeleccionar.AppendLine("INTO	#tmpTemporal")
            loComandoSeleccionar.AppendLine("FROM	#tmpCambios")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = #tmpCambios.Codigo")
            loComandoSeleccionar.AppendLine("WHERE	Detalle.exist('(/detalle/campos/campo[@nombre=""precio1""]/despues)[1]')= 1")
            loComandoSeleccionar.AppendLine("	OR	Detalle.exist('(/detalle/campos/campo[@nombre=""fec_pre1""]/despues)[1]')= 1")
            loComandoSeleccionar.AppendLine("	OR	Detalle.exist('(/detalle/campos/campo[@nombre=""fec_pre2""]/despues)[1]')= 1")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpCambios.Codigo, #tmpCambios.Registro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCambios")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Usu, Registro, Nom_cli, Rif_Cli, Cod_Art, ")
            loComandoSeleccionar.AppendLine("		ISNULL(Canon_Anterior, ")
            loComandoSeleccionar.AppendLine("			(	SELECT	TOP 1 T.Canon_Nuevo ")
            loComandoSeleccionar.AppendLine("				FROM	#tmpTemporal AS T ")
            loComandoSeleccionar.AppendLine("				WHERE	(T.Registro < G.Registro) ")
            loComandoSeleccionar.AppendLine("					AND (NOT T.Canon_Nuevo IS NULL)")
            loComandoSeleccionar.AppendLine("					AND (T.Cod_Art = G.Cod_Art)")
            loComandoSeleccionar.AppendLine("				ORDER BY T.Registro DESC))	AS Canon_Anterior, ")
            loComandoSeleccionar.AppendLine("		ISNULL(Canon_Nuevo, ")
            loComandoSeleccionar.AppendLine("			(	SELECT	TOP 1 T.Canon_Nuevo ")
            loComandoSeleccionar.AppendLine("				FROM	#tmpTemporal AS T ")
            loComandoSeleccionar.AppendLine("				WHERE	(T.Registro < G.Registro) ")
            loComandoSeleccionar.AppendLine("					AND (NOT T.Canon_Nuevo IS NULL)")
            loComandoSeleccionar.AppendLine("					AND (T.Cod_Art = G.Cod_Art)")
            loComandoSeleccionar.AppendLine("				ORDER BY T.Registro DESC))	AS Canon_Nuevo,")
            loComandoSeleccionar.AppendLine("		ISNULL(Fecha_Inicio, ")
            loComandoSeleccionar.AppendLine("			(	SELECT	TOP 1 T.Fecha_Inicio ")
            loComandoSeleccionar.AppendLine("				FROM	#tmpTemporal AS T ")
            loComandoSeleccionar.AppendLine("				WHERE	(T.Registro < G.Registro) ")
            loComandoSeleccionar.AppendLine("					AND (NOT T.Fecha_Inicio IS NULL)")
            loComandoSeleccionar.AppendLine("					AND (T.Cod_Art = G.Cod_Art)")
            loComandoSeleccionar.AppendLine("				ORDER BY T.Registro DESC))	AS Fecha_Inicio,")
            loComandoSeleccionar.AppendLine("		ISNULL(Fecha_Vencimiento, ")
            loComandoSeleccionar.AppendLine("			(	SELECT	TOP 1 T.Fecha_Vencimiento ")
            loComandoSeleccionar.AppendLine("				FROM	#tmpTemporal AS T ")
            loComandoSeleccionar.AppendLine("				WHERE	(T.Registro < G.Registro) ")
            loComandoSeleccionar.AppendLine("					AND (NOT T.Fecha_Vencimiento IS NULL)")
            loComandoSeleccionar.AppendLine("					AND (T.Cod_Art = G.Cod_Art)")
            loComandoSeleccionar.AppendLine("				ORDER BY T.Registro DESC))	AS Fecha_Vencimiento")
            loComandoSeleccionar.AppendLine("FROM	#tmpTemporal AS G")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rHistorico_Canon", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrHistorico_Canon.ReportSource = loObjetoReporte

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
' RJG: 15/08/11: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 22/08/11: Ajuste para rrellenar valores vacíos en el reporte.						'
'-------------------------------------------------------------------------------------------'
