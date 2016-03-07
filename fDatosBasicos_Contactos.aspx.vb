'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDatosBasicos_Contactos"
'-------------------------------------------------------------------------------------------'
Partial Class fDatosBasicos_Contactos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
			Dim lcTipoContactos as String = goEmpresa.pcCodigo

			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Contactos.cod_con,")
			loComandoSeleccionar.AppendLine(" 			Contactos.nom_con,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Contactos.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Contactos.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Contactos.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Contactos.sexo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.cedula,")
			loComandoSeleccionar.AppendLine(" 			Contactos.empresa,")
			loComandoSeleccionar.AppendLine(" 			Contactos.dep_con,")
			loComandoSeleccionar.AppendLine(" 			Contactos.oficina,")
			loComandoSeleccionar.AppendLine(" 			Contactos.directo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.personal,")
			loComandoSeleccionar.AppendLine(" 			Contactos.telefonos,")
			loComandoSeleccionar.AppendLine(" 			Contactos.fax,")
			loComandoSeleccionar.AppendLine(" 			Contactos.extension,")
			loComandoSeleccionar.AppendLine(" 			Contactos.movil,")
			loComandoSeleccionar.AppendLine(" 			Contactos.correo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.adicional,")
			loComandoSeleccionar.AppendLine(" 			Contactos.direccion,")
			loComandoSeleccionar.AppendLine(" 			Contactos.titulo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.cargo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.niv_dec,")
			loComandoSeleccionar.AppendLine(" 			Contactos.web,")
			loComandoSeleccionar.AppendLine(" 			Contactos.clave,")
			loComandoSeleccionar.AppendLine(" 			Contactos.fec_ini,")
			loComandoSeleccionar.AppendLine(" 			Contactos.fec_fin,")
			loComandoSeleccionar.AppendLine(" 			Contactos.fec_nac,")
			loComandoSeleccionar.AppendLine(" 			Contactos.comentario,")
			loComandoSeleccionar.AppendLine(" 			Contactos.defecto,")
			loComandoSeleccionar.AppendLine(" 			Contactos.tipo,")
			loComandoSeleccionar.AppendLine(" 			Contactos.cod_reg")
			loComandoSeleccionar.AppendLine(" INTO	#tmpTemporal		")
			loComandoSeleccionar.AppendLine(" FROM Contactos")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)

			loComandoSeleccionar.AppendLine(" 	DECLARE	@lcTipo VARCHAR (15)	")
			loComandoSeleccionar.AppendLine(" 	SET @lcTipo = (SELECT Tipo 	")
			loComandoSeleccionar.AppendLine(" 				FROM #tmpTemporal) 	")
			
			loComandoSeleccionar.AppendLine(" 	WHILE @lcTipo ='" & lcTipoContactos & "Prospectos'")
			loComandoSeleccionar.AppendLine(" 	BEGIN	")
						loComandoSeleccionar.AppendLine(" 	SELECT 		")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.nom_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.status,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.sexo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cedula,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.empresa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.dep_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.oficina,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.directo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.personal,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.telefonos,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fax,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.extension,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.movil,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.correo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.adicional,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.direccion,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.titulo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cargo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.niv_dec,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.web,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.clave,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_ini,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_fin,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_nac,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.comentario,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.defecto,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.tipo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_reg")
						loComandoSeleccionar.AppendLine(" FROM	#tmpTemporal,Prospectos		")
						loComandoSeleccionar.AppendLine(" WHERE #tmpTemporal.Cod_Reg = Prospectos.Cod_Pro")
			loComandoSeleccionar.AppendLine(" Break")
			loComandoSeleccionar.AppendLine(" End ")
					
			loComandoSeleccionar.AppendLine(" 	WHILE @lcTipo ='" & lcTipoContactos & "Clientes'")
			loComandoSeleccionar.AppendLine(" 	BEGIN	")
						loComandoSeleccionar.AppendLine(" 	SELECT 		")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.nom_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.status,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.sexo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cedula,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.empresa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.dep_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.oficina,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.directo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.personal,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.telefonos,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fax,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.extension,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.movil,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.correo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.adicional,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.direccion,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.titulo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cargo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.niv_dec,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.web,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.clave,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_ini,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_fin,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_nac,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.comentario,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.defecto,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.tipo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_reg")
						loComandoSeleccionar.AppendLine(" 	FROM	#tmpTemporal,Clientes		")
						loComandoSeleccionar.AppendLine(" WHERE #tmpTemporal.Cod_Reg  = Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine(" Break")
			loComandoSeleccionar.AppendLine(" End ")
			
			loComandoSeleccionar.AppendLine(" 	WHILE @lcTipo ='" & lcTipoContactos & "Usuarios'")
			loComandoSeleccionar.AppendLine(" 	BEGIN	")
						loComandoSeleccionar.AppendLine(" 	SELECT 		")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.nom_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.status,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.sexo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cedula,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.empresa,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.dep_con,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.oficina,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.directo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.personal,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.telefonos,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fax,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.extension,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.movil,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.correo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.adicional,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.direccion,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.titulo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cargo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.niv_dec,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.web,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.clave,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_ini,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_fin,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.fec_nac,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.comentario,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.defecto,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.tipo,")
						loComandoSeleccionar.AppendLine(" 			#tmpTemporal.cod_reg")
						loComandoSeleccionar.AppendLine(" 	FROM	#tmpTemporal,Contactos		")
						loComandoSeleccionar.AppendLine(" WHERE #tmpTemporal.Cod_Reg  = Contactos.Cod_Usu and #tmpTemporal.Cod_Con  = Contactos.Cod_Con ")
			loComandoSeleccionar.AppendLine(" Break")
			loComandoSeleccionar.AppendLine(" End ")
			
			'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDatosBasicos_Contactos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDatosBasicos_Contactos.ReportSource = loObjetoReporte

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
' MAT: 20/12/10: Codigo inicial
'-------------------------------------------------------------------------------------------'