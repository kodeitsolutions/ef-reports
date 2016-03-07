'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fEventos"
'-------------------------------------------------------------------------------------------'
Partial Class fEventos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 	cod_eve,")				
			loComandoSeleccionar.AppendLine(" 	nom_eve,")	
			loComandoSeleccionar.AppendLine(" 	CASE")
			loComandoSeleccionar.AppendLine(" 		WHEN status = 'A' THEN 'Acttivos'")
			loComandoSeleccionar.AppendLine(" 		WHEN status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 		WHEN status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 	END AS status,")			
			loComandoSeleccionar.AppendLine(" 	nivel,")				
			loComandoSeleccionar.AppendLine(" 	fec_ini,")				
			loComandoSeleccionar.AppendLine(" 	tip_dis,")				
			loComandoSeleccionar.AppendLine(" 	comentario,")				
			loComandoSeleccionar.AppendLine(" 	notas,")				
			loComandoSeleccionar.AppendLine(" 	dat_eje,")				
			loComandoSeleccionar.AppendLine(" 	eje_cua,")				
			loComandoSeleccionar.AppendLine(" 	cor_par,")				
			loComandoSeleccionar.AppendLine(" 	cor_cop,")				
			loComandoSeleccionar.AppendLine(" 	cor_ocu,")				
			loComandoSeleccionar.AppendLine(" 	cor_asu,")				
			loComandoSeleccionar.AppendLine(" 	cor_men,")				
			loComandoSeleccionar.AppendLine(" 	cor_adj,")				
			loComandoSeleccionar.AppendLine(" 	usuarios,")				
			loComandoSeleccionar.AppendLine(" 	tip_not,")				
			loComandoSeleccionar.AppendLine(" 	prioridad,")				
			loComandoSeleccionar.AppendLine(" 	titulo,")				
			loComandoSeleccionar.AppendLine(" 	contenido,")				
			loComandoSeleccionar.AppendLine(" 	sen_sql,")				
			loComandoSeleccionar.AppendLine(" 	modulo,")				
			loComandoSeleccionar.AppendLine(" 	seccion,")				
			loComandoSeleccionar.AppendLine(" 	opcion,")				
			loComandoSeleccionar.AppendLine(" 	gancho,")				
			loComandoSeleccionar.AppendLine(" 	con_ope,")				
			loComandoSeleccionar.AppendLine(" 	tit_men,")				
			loComandoSeleccionar.AppendLine(" 	con_men,")				
			loComandoSeleccionar.AppendLine(" 	tip_eje,")				
			loComandoSeleccionar.AppendLine(" 	tip_rep,")				
			loComandoSeleccionar.AppendLine(" 	fin_dia,")				
			loComandoSeleccionar.AppendLine(" 	fin_rep,")				
			loComandoSeleccionar.AppendLine(" 	rep_dia,")				
			loComandoSeleccionar.AppendLine(" 	eje_sem,")				
			loComandoSeleccionar.AppendLine(" 	dia_sem,")				
			loComandoSeleccionar.AppendLine(" 	rep_men,")				
			loComandoSeleccionar.AppendLine(" 	dur_dia,")				
			loComandoSeleccionar.AppendLine(" 	dia_eje,")				
			loComandoSeleccionar.AppendLine(" 	mes_sem,")				
			loComandoSeleccionar.AppendLine(" 	mes_dia,")				
			loComandoSeleccionar.AppendLine(" 	anu_dia,")				
			loComandoSeleccionar.AppendLine(" 	dia_anu,")				
			loComandoSeleccionar.AppendLine(" 	mes_anu,")				
			loComandoSeleccionar.AppendLine(" 	anu_dia2,")				
			loComandoSeleccionar.AppendLine(" 	anu_dia3,")				
			loComandoSeleccionar.AppendLine(" 	mes_dia4,")				
			loComandoSeleccionar.AppendLine(" 	tip_fin,")				
			loComandoSeleccionar.AppendLine(" 	eje_mes,")				
			loComandoSeleccionar.AppendLine(" 	eje_año")				
			loComandoSeleccionar.AppendLine(" FROM eventos")			

			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fEventos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)	   
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfEventos.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 22/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 11/05/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'