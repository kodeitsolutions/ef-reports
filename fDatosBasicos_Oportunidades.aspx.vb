'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatosBasicos_Oportunidades"
'-------------------------------------------------------------------------------------------'
Partial Class fDatosBasicos_Oportunidades
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

		
            Dim loComandoSeleccionar As New StringBuilder()
			

			loComandoSeleccionar.AppendLine(" 	SELECT 		")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Origen,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Reg,")		
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Opo,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Nom_Opo,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Procedencia,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Pro_Pas,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Fec_Fin,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Status,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Ven,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Mon,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Probabilidad,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Mon_Net,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Tasa,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Prioridad,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Usu,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cod_Gru,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Cantidad,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Por_Eje,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Etapa,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Tip_Pro,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Tip_Opo,")
			loComandoSeleccionar.AppendLine(" 			Oportunidades.Comentario")
			loComandoSeleccionar.AppendLine(" 	INTO	#tmpTemporal		")
			loComandoSeleccionar.AppendLine(" 	FROM	Oportunidades		")
			loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			loComandoSeleccionar.AppendLine(" 	DECLARE	@lcOrigen VARCHAR (15)	")
			loComandoSeleccionar.AppendLine(" 	SET @lcOrigen = (SELECT Origen 	")
			loComandoSeleccionar.AppendLine(" 				FROM #tmpTemporal) 	")
			
			loComandoSeleccionar.AppendLine(" 	WHILE @lcOrigen = 'Prospectos'")
			loComandoSeleccionar.AppendLine(" 	BEGIN	")
						loComandoSeleccionar.AppendLine(" 	SELECT 		")
						loComandoSeleccionar.AppendLine(" 			Prospectos.Cod_Pro As Codigo,")
						loComandoSeleccionar.AppendLine(" 			Prospectos.Nom_Pro As Nombre,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Origen,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Reg,")		
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Nom_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Procedencia,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Pro_Pas,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Fec_Ini,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Fec_Fin,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Status,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Ven,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Mon,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Probabilidad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Mon_Net,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tasa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Prioridad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Usu,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Gru,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cantidad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Por_Eje,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Etapa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tip_Pro,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tip_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Comentario")
						loComandoSeleccionar.AppendLine(" 	FROM	#tmpTemporal,Prospectos		")
						loComandoSeleccionar.AppendLine(" WHERE #tmpTemporal.Cod_Reg = Prospectos.Cod_Pro")
			loComandoSeleccionar.AppendLine(" Break")
			loComandoSeleccionar.AppendLine(" End ")
			
		
			
			loComandoSeleccionar.AppendLine(" 	WHILE @lcOrigen = 'Clientes'")
			loComandoSeleccionar.AppendLine(" 	BEGIN	")
						loComandoSeleccionar.AppendLine(" 	SELECT 		")
						loComandoSeleccionar.AppendLine(" 			Clientes.Cod_Cli As Codigo,")
						loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli As Nombre,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Origen,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Reg,")		
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Nom_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Procedencia,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Pro_Pas,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Fec_Ini,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Fec_Fin,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Status,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Ven,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Mon,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Probabilidad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Mon_Net,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tasa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Prioridad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Usu,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cod_Gru,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Cantidad,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Por_Eje,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Etapa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tip_Pro,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Tip_Opo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.Comentario")
						loComandoSeleccionar.AppendLine(" 	FROM	#tmpTemporal,Clientes		")
						loComandoSeleccionar.AppendLine(" WHERE #tmpTemporal.Cod_Reg  = Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine(" Break")
			loComandoSeleccionar.AppendLine(" End ")
			
			'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

           	'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
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
					  
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatosBasicos_Oportunidades", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatosBasicos_Oportunidades.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
'MAT:  20/12/10 : Codigo inicial															'
'-------------------------------------------------------------------------------------------'

