'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrabajadores_CNomina"
'-------------------------------------------------------------------------------------------'
Partial Class rTrabajadores_CNomina
Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try
		
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		Campos_Nomina.Cod_Cam				AS Cod_Cam,")
			loComandoSeleccionar.AppendLine("			Campos_Nomina.Nom_Cam				AS Nom_Cam,")
			loComandoSeleccionar.AppendLine("			Campos_Nomina.Status				AS Status,")
			loComandoSeleccionar.AppendLine("			Campos_Nomina.Comentario			AS Comentario,")
			loComandoSeleccionar.AppendLine("			(CASE WHEN Campos_Nomina.Status = 'A'")
			loComandoSeleccionar.AppendLine("				THEN 'Activo'			")
			loComandoSeleccionar.AppendLine("				ELSE 'Inactivo'			")
			loComandoSeleccionar.AppendLine("			END)								AS Estatus,	 ")
			loComandoSeleccionar.AppendLine("			(CASE Campos_Nomina.Tipo	")
			loComandoSeleccionar.AppendLine("				WHEN 'C' THEN 'Carácter'")
			loComandoSeleccionar.AppendLine("				WHEN 'N' THEN 'Numérico'")
			loComandoSeleccionar.AppendLine("				WHEN 'L' THEN 'Lógico'	")
			loComandoSeleccionar.AppendLine("				WHEN 'F' THEN 'Fecha'	")
			loComandoSeleccionar.AppendLine("				WHEN 'M' THEN 'Memo'	")
			loComandoSeleccionar.AppendLine("				ELSE '[Desconocido]'	")
			loComandoSeleccionar.AppendLine("			END)								AS Tipo,")
			loComandoSeleccionar.AppendLine("			Campos_Nomina.Uso					AS Uso,")
			loComandoSeleccionar.AppendLine("			Trabajadores.cod_tra				AS Cod_Tra,")
			loComandoSeleccionar.AppendLine("			Trabajadores.nom_tra				AS Nom_Tra,")
			loComandoSeleccionar.AppendLine("			renglones_Campos_Nomina.Val_car		AS Val_car,")
			loComandoSeleccionar.AppendLine("			renglones_Campos_Nomina.Val_Num		AS Val_Num,")
			loComandoSeleccionar.AppendLine("			renglones_Campos_Nomina.Val_Fec 	AS Val_Fec,")
			loComandoSeleccionar.AppendLine("			renglones_Campos_Nomina.Val_Log		AS Val_Log,")
			loComandoSeleccionar.AppendLine("			renglones_Campos_Nomina.Val_Tex		AS Val_Tex")
			loComandoSeleccionar.AppendLine("			")
			loComandoSeleccionar.AppendLine("FROM		Campos_Nomina  ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN renglones_Campos_Nomina  ")
			loComandoSeleccionar.AppendLine("		ON	renglones_Campos_Nomina.cod_cam = Campos_Nomina.cod_cam")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Trabajadores  ")
			loComandoSeleccionar.AppendLine("		ON	Trabajadores.cod_tra = renglones_Campos_Nomina.cod_tra")
			loComandoSeleccionar.AppendLine("WHERE	Trabajadores.cod_tra BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Trabajadores.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("		AND	Campos_Nomina.Cod_Cam BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("		AND	" & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND	Campos_Nomina.Tipo IN (" & lcParametro3Desde & ")")
			loComandoSeleccionar.AppendLine("		AND	Campos_Nomina.Uso IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrabajadores_CNomina", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrTrabajadores_CNomina.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/03/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
