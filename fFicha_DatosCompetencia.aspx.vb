'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosCompetencia"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosCompetencia
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("  SELECT ")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_com, ")
			loComandoSeleccionar.AppendLine(" 			Competencia.nom_com, ")
			loComandoSeleccionar.AppendLine(" 			 			CASE")
			loComandoSeleccionar.AppendLine(" 							WHEN Competencia.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 							WHEN Competencia.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 							WHEN Competencia.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 						END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Competencia.rif,")
			loComandoSeleccionar.AppendLine(" 			Competencia.nit,")
			loComandoSeleccionar.AppendLine(" 			Competencia.dir_fis,")
			loComandoSeleccionar.AppendLine(" 			Competencia.dir_exa,")
			loComandoSeleccionar.AppendLine(" 			Competencia.dir_otr,")
			loComandoSeleccionar.AppendLine(" 			Competencia.telefonos,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.directo,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.fax,")
			loComandoSeleccionar.AppendLine(" 			Competencia.movil,		")
			loComandoSeleccionar.AppendLine(" 			Competencia.correo,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.correo2,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.correo3,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.web,")
			loComandoSeleccionar.AppendLine(" 			Competencia.kilometros,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_pai,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_est,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_ciu,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_zon,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_ven,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_for,")
			loComandoSeleccionar.AppendLine(" 			Competencia.cod_tra,")
			loComandoSeleccionar.AppendLine(" 			Competencia.tip_pag,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.tip_emp,")
			loComandoSeleccionar.AppendLine(" 			Competencia.abc,")
			loComandoSeleccionar.AppendLine(" 			Competencia.mar_rec,")
			loComandoSeleccionar.AppendLine(" 			Competencia.can_emp,")
			loComandoSeleccionar.AppendLine(" 			Competencia.ven_men,")
			loComandoSeleccionar.AppendLine(" 			Competencia.hor_caj,")
			loComandoSeleccionar.AppendLine(" 			Competencia.pos_mer,")
			loComandoSeleccionar.AppendLine(" 			Competencia.perfil,")
			loComandoSeleccionar.AppendLine(" 			Competencia.pun_fue,")
			loComandoSeleccionar.AppendLine(" 			Competencia.pun_deb,")
			loComandoSeleccionar.AppendLine(" 			Competencia.clientes,")
			loComandoSeleccionar.AppendLine(" 			Competencia.det_cli,")
			loComandoSeleccionar.AppendLine(" 			Competencia.comentario,")
			loComandoSeleccionar.AppendLine(" 			Competencia.tipo,")
			loComandoSeleccionar.AppendLine(" 			Competencia.clase,")
			loComandoSeleccionar.AppendLine(" 			Competencia.grupo,")
			loComandoSeleccionar.AppendLine(" 			Competencia.valor1,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.valor2,")
			loComandoSeleccionar.AppendLine(" 			Competencia.valor3,")
			loComandoSeleccionar.AppendLine(" 			Competencia.valor4,")
			loComandoSeleccionar.AppendLine(" 			Competencia.valor5,	")
			loComandoSeleccionar.AppendLine(" 			Competencia.nivel,")
			loComandoSeleccionar.AppendLine(" 			Competencia.prioridad,")
			loComandoSeleccionar.AppendLine(" 			Competencia.contacto,")
			loComandoSeleccionar.AppendLine("  			Zonas.Nom_Zon,")
			loComandoSeleccionar.AppendLine("  			Paises.Nom_Pai,")
			loComandoSeleccionar.AppendLine("  			Estados.Nom_Est,")
			loComandoSeleccionar.AppendLine("  			Ciudades.Nom_Ciu,")
			loComandoSeleccionar.AppendLine("  			Formas_Pagos.Nom_For,")
			loComandoSeleccionar.AppendLine("  			Vendedores.Nom_Ven,")
			loComandoSeleccionar.AppendLine("  			Transportes.Nom_Tra")			
			loComandoSeleccionar.AppendLine(" FROM Competencia")
			loComandoSeleccionar.AppendLine(" JOIN Zonas ON Zonas.Cod_Zon = Competencia.Cod_Zon")
			loComandoSeleccionar.AppendLine(" JOIN Paises ON Paises.Cod_Pai = Competencia.Cod_Pai")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Estados ON Estados.Cod_Est = Competencia.Cod_Est")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Ciudades ON Ciudades.Cod_Ciu = Competencia.Cod_Ciu")
			loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Competencia.Cod_For")
			loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Competencia.Cod_Ven")
			loComandoSeleccionar.AppendLine(" JOIN Transportes ON Transportes.Cod_Tra = Competencia.Cod_Tra")
			loComandoSeleccionar.AppendLine(" WHERE ")	   			
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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosCompetencia", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosCompetencia.ReportSource = loObjetoReporte

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
' CMS: 21/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 11/05/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'