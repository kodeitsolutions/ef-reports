'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosProspectos"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosProspectos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
			Dim lcTipo As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo & "Prospectos")

			loComandoSeleccionar.AppendLine("  SELECT	nacional,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_Pro,")
			loComandoSeleccionar.AppendLine("  			Prospectos.nom_Pro,")
			loComandoSeleccionar.AppendLine("  			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine("  			Prospectos.tip_Cli,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.crm_sec = 'P' THEN 'Producción'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.crm_sec = 'C' THEN 'Comercio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.crm_sec = 'S' THEN 'Servicio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.crm_sec = 'G' THEN 'Gobierno'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.crm_sec = 'O' THEN 'Otro'")
			loComandoSeleccionar.AppendLine(" 			END AS crm_sec,")
			loComandoSeleccionar.AppendLine("  			Prospectos.abc,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Sexo = 'X' THEN 'No Aplica'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Sexo = 'M' THEN 'Masculino'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.Sexo = 'F' THEN 'Femenino'")
			loComandoSeleccionar.AppendLine(" 			END AS Sexo,")
			loComandoSeleccionar.AppendLine("  			Prospectos.tip_pag,")
			loComandoSeleccionar.AppendLine("  			Prospectos.nit,")
			loComandoSeleccionar.AppendLine("  			Prospectos.rif,")
			loComandoSeleccionar.AppendLine("  			Prospectos.movil,")
			loComandoSeleccionar.AppendLine("  			Prospectos.web,")
			loComandoSeleccionar.AppendLine("  			Prospectos.telefonos,")
			loComandoSeleccionar.AppendLine("  			Prospectos.fec_ini,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_pos,")
			loComandoSeleccionar.AppendLine("  			Prospectos.fax,")
			loComandoSeleccionar.AppendLine("  			Prospectos.fec_fin,")
			loComandoSeleccionar.AppendLine("  			Prospectos.correo,")
			loComandoSeleccionar.AppendLine("  			Prospectos.correo2,")
			loComandoSeleccionar.AppendLine("  			Prospectos.correo3,")
			loComandoSeleccionar.AppendLine("  			Prospectos.contacto,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_tip,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_zon,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_pai,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_est,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_ciu,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_cla,")
			loComandoSeleccionar.AppendLine(" 			Prospectos.cod_con,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_per,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_for,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_ven,")
			loComandoSeleccionar.AppendLine("  			Prospectos.cod_tra,")
			loComandoSeleccionar.AppendLine("  			Prospectos.matriz,")
			loComandoSeleccionar.AppendLine("  			Prospectos.dir_fis,")
			loComandoSeleccionar.AppendLine("  			Prospectos.con_esp,")
			loComandoSeleccionar.AppendLine("  			Prospectos.dir_exa,")
			loComandoSeleccionar.AppendLine("  			Prospectos.dir_otr,")
			loComandoSeleccionar.AppendLine("  			Prospectos.dir_ent,")
			loComandoSeleccionar.AppendLine("  			Prospectos.actividad,")
			loComandoSeleccionar.AppendLine("  			Prospectos.are_neg,")
			loComandoSeleccionar.AppendLine("  			Prospectos.fec_nac,")
			loComandoSeleccionar.AppendLine("  			Prospectos.comentario,")
			loComandoSeleccionar.AppendLine("  			Prospectos.notas,")
			loComandoSeleccionar.AppendLine("  			Prospectos.hor_caj,")
			loComandoSeleccionar.AppendLine("  			Prospectos.fiscal,")
			loComandoSeleccionar.AppendLine("  			Prospectos.nacional,")
			loComandoSeleccionar.AppendLine("  			Prospectos.mar_rec,")
			loComandoSeleccionar.AppendLine("  			Prospectos.crm_pos,")
			loComandoSeleccionar.AppendLine("  			Prospectos.pol_com,")
			loComandoSeleccionar.AppendLine("  			Prospectos.mon_sal,")
			loComandoSeleccionar.AppendLine("  			Prospectos.mon_cre,")
			loComandoSeleccionar.AppendLine("  			Prospectos.dia_cre,")
			loComandoSeleccionar.AppendLine("  			Prospectos.pro_pag,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_des,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_ali,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_int,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_pat,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_isl,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_ret,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_pat,")
			loComandoSeleccionar.AppendLine("  			Prospectos.por_otr,")
			loComandoSeleccionar.AppendLine("  			Prospectos.kilometros,")
			loComandoSeleccionar.AppendLine("  			Prospectos.generico,")
			loComandoSeleccionar.AppendLine("  			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Formal' THEN 'Formal'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Ordinario' THEN 'Ordinario'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Especial' THEN 'Especial'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Ocasional' THEN 'Ocasional'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Otro_1' THEN 'Otro 1'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Otro_2' THEN 'Otro 2'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Otro_3' THEN 'Otro 3'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Otro_4' THEN 'Otro 4'")
			loComandoSeleccionar.AppendLine(" 				WHEN Prospectos.tip_con = 'Otro_5' THEN 'Otro 5'")
			loComandoSeleccionar.AppendLine(" 			END AS tip_con,")
			loComandoSeleccionar.AppendLine("  			Prospectos.tipo,")
			loComandoSeleccionar.AppendLine("  			Prospectos.clase,")
			loComandoSeleccionar.AppendLine("  			Tipos_Clientes.Nom_Tip,")
			loComandoSeleccionar.AppendLine("  			Zonas.Nom_Zon,")
			loComandoSeleccionar.AppendLine("  			Paises.Nom_Pai,")
			loComandoSeleccionar.AppendLine("  			Estados.Nom_Est,")
			loComandoSeleccionar.AppendLine("  			Ciudades.Nom_Ciu,")
			loComandoSeleccionar.AppendLine("  			Clases_Clientes.Nom_Cla,")
			loComandoSeleccionar.AppendLine("  			Personas.Nom_Per,")
			loComandoSeleccionar.AppendLine("  			Formas_Pagos.Nom_For,")
			loComandoSeleccionar.AppendLine("  			Vendedores.Nom_Ven,")
			loComandoSeleccionar.AppendLine("  			Transportes.Nom_Tra,")
			loComandoSeleccionar.AppendLine("  			Conceptos.Nom_Con,")
			loComandoSeleccionar.AppendLine(" 			Contactos.Nom_Con	As Nom_Contacto,")
			loComandoSeleccionar.AppendLine(" 			Contactos.Cargo	As Cargo_Con,")
			loComandoSeleccionar.AppendLine(" 			Contactos.Telefonos As Tlf_Con,")
			loComandoSeleccionar.AppendLine(" 			Contactos.Correo As Mail_Con")
			loComandoSeleccionar.AppendLine(" FROM Prospectos")
			loComandoSeleccionar.AppendLine(" JOIN Tipos_Clientes ON Tipos_Clientes.Cod_Tip = Prospectos.Cod_Tip")
			loComandoSeleccionar.AppendLine(" JOIN Zonas ON Zonas.Cod_Zon = Prospectos.Cod_Zon")
			loComandoSeleccionar.AppendLine(" JOIN Paises ON Paises.Cod_Pai = Prospectos.Cod_Pai")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Estados ON Estados.Cod_Est = Prospectos.Cod_Est")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Ciudades ON Ciudades.Cod_Ciu = Prospectos.Cod_Ciu")
			loComandoSeleccionar.AppendLine(" JOIN Clases_Clientes ON Clases_Clientes.Cod_Cla = Prospectos.Cod_Cla  ")
			loComandoSeleccionar.AppendLine(" JOIN Personas ON Personas.Cod_Per = Prospectos.Cod_Per")
			loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Prospectos.Cod_For")
			loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Prospectos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" JOIN Transportes ON Transportes.Cod_Tra = Prospectos.Cod_Tra")
			loComandoSeleccionar.AppendLine(" JOIN Conceptos ON Conceptos.Cod_Con = Prospectos.Cod_Con")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Contactos ON Contactos.Tipo =" & lcTipo & " AND Contactos.Cod_Reg = Prospectos.Cod_Pro")
			loComandoSeleccionar.AppendLine(" WHERE Prospectos.Origen = 'Prospectos'")	 	   	 			
            loComandoSeleccionar.AppendLine("          AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosProspectos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosProspectos.ReportSource = loObjetoReporte

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
' CMS: 19/09/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 11/05/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 22/06/2011 : Ajuste en el Select
'-------------------------------------------------------------------------------------------'