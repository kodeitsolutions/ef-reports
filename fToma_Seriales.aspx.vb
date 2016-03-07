'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fToma_Seriales"
'-------------------------------------------------------------------------------------------'
Partial Class fToma_Seriales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("-- Lista de seriales: uno por renglón")
            loComandoSeleccionar.AppendLine("SELECT		Tomas.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("			Tomas.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			CAST(Tomas.Comentario AS VARCHAR(1000))	AS Comentario,")
            loComandoSeleccionar.AppendLine("			CAST(Tomas.Referencia AS VARCHAR(100))	AS Referencia,")
            loComandoSeleccionar.AppendLine("			Renglones_Tomas.Renglon					AS Renglon,")
            loComandoSeleccionar.AppendLine("			Renglones_Tomas.Cod_Art					AS Cod_Art,")
            loComandoSeleccionar.AppendLine("			Renglones_Tomas.Nom_Art					AS Nom_Art,")
            loComandoSeleccionar.AppendLine("			Renglones_Tomas.Serial					AS Serial,")
            loComandoSeleccionar.AppendLine("INTO		#tmpSeriales")
            loComandoSeleccionar.AppendLine("FROM		Tomas")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Tomas")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Tomas.Documento = Tomas.Documento")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("ORDER BY	Renglones_Tomas.Cod_Art, Renglones_Tomas.Serial")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Lista de seriales, separados por coma")
            loComandoSeleccionar.AppendLine("SELECT		Externo.Documento											AS Documento,  ")
            loComandoSeleccionar.AppendLine("			Externo.Fec_Ini												AS Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("			Externo.Comentario											AS Comentario, ")
            loComandoSeleccionar.AppendLine("			Externo.Referencia											AS Referencia, ")
            loComandoSeleccionar.AppendLine("			Externo.Cod_Art												AS Cod_Art,  ")
            loComandoSeleccionar.AppendLine("			Externo.Nom_Art												AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("			(	SUBSTRING(CAST(	")
            loComandoSeleccionar.AppendLine("					(	SELECT	', ' + RTRIM(Interno.Serial)  ")
            loComandoSeleccionar.AppendLine("						FROM	#tmpSeriales AS Interno ")
            loComandoSeleccionar.AppendLine("						WHERE	Interno.Cod_Art = Externo.Cod_Art ")
            loComandoSeleccionar.AppendLine("						ORDER BY Interno.Serial FOR XML PATH('') ")
            loComandoSeleccionar.AppendLine("					) AS VARCHAR(MAX)), 3, 1000000000)")
            loComandoSeleccionar.AppendLine("			)																AS Seriales")
            loComandoSeleccionar.AppendLine("FROM		#tmpSeriales AS Externo ")
            loComandoSeleccionar.AppendLine("GROUP BY	Externo.Documento,  ")
            loComandoSeleccionar.AppendLine("			Externo.Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("			Externo.Comentario, ")
            loComandoSeleccionar.AppendLine("			Externo.Referencia, ")
            loComandoSeleccionar.AppendLine("			Externo.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Externo.Nom_Art")
            loComandoSeleccionar.AppendLine("			  ")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSeriales")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine()

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

   			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fToma_Seriales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfToma_Seriales.ReportSource = loObjetoReporte

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
' RJG: 02/11/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
