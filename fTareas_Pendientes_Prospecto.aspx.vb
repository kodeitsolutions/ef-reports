'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTareas_Pendientes_Prospecto"
'-------------------------------------------------------------------------------------------'
Partial Class fTareas_Pendientes_Prospecto
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
      
			loComandoSeleccionar.AppendLine ("declare @lcProspecto char(30)")
			loComandoSeleccionar.AppendLine("set @lcProspecto = (select top 1 cod_reg from tareas where " & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" select ")
			loComandoSeleccionar.AppendLine(" 		Tareas.Origen,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Cod_Reg,")
			loComandoSeleccionar.AppendLine(" 		Prospectos.Cod_Pro Nombre,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Cod_Tar,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Nom_Tar,")
			loComandoSeleccionar.AppendLine(" 		(CASE WHEN Tareas.Tipo = '' THEN '[SIN TIPO ASIGNADO]' ELSE Tareas.Tipo END) Tipo,")
			loComandoSeleccionar.AppendLine(" 		(CASE WHEN Tareas.Clase = '' THEN '[SIN CLASE ASIGNADA]' ELSE Tareas.Clase END) Clase,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Comentario,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Notas,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Importancia,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Prioridad,")
			loComandoSeleccionar.AppendLine(" 		CASE")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'No Iniciada'		THEN 'No Iniciada'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'En Progreso'		THEN 'En Progreso'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Completada'		THEN 'Completada'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Pendiente Datos'	THEN 'Pendiente por Datos'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Retrasada'			THEN 'Retrasada'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Planeada'			THEN 'Planeada'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Otros1'			THEN 'Otros 1'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Otros2'			THEN 'Otros 2'")
			loComandoSeleccionar.AppendLine(" 			WHEN Tareas.Etapa = 'Otros3'			THEN 'Otros 3'")
			loComandoSeleccionar.AppendLine(" 		END AS Etapa,")
			loComandoSeleccionar.AppendLine(" 		Tareas.Cod_Gru")
			loComandoSeleccionar.AppendLine("FROM   Tareas")
			loComandoSeleccionar.AppendLine("join Prospectos on Prospectos.Cod_Pro = Tareas.Cod_Reg")
            loComandoSeleccionar.AppendLine("WHERE	 Tareas.Etapa <> 'Completada'")
            loComandoSeleccionar.AppendLine("	 AND Tareas.Origen = 'Prospectos'")
            loComandoSeleccionar.AppendLine("	 AND Tareas.Cod_Reg = @lcProspecto" )
            loComandoSeleccionar.AppendLine("ORDER BY Tareas.Tipo ASC, Tareas.Clase ASC")
            'loComandoSeleccionar.AppendLine("    AND " & cusAplicacion.goFormatos.pcCondicionPrincipal & " ")


		    'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTareas_Pendientes_Prospecto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfTareas_Pendientes_Prospecto.ReportSource = loObjetoReporte

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
' CMS: 20/07/2010: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' MAT: 17/01/2011: Adición campos Nuevos, según requerimientos
'-------------------------------------------------------------------------------------------'