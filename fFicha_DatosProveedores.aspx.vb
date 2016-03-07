'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosProveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_pro,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.nom_pro,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.tip_pro,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.crm_sec = 'P' THEN 'Producción'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.crm_sec = 'C' THEN 'Comercio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.crm_sec = 'S' THEN 'Servicio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.crm_sec = 'G' THEN 'Gobierno'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.crm_sec = 'O' THEN 'Otro'")
			loComandoSeleccionar.AppendLine(" 			END AS crm_sec,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.abc,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Sexo = 'X' THEN 'No Aplica'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Sexo = 'M' THEN 'Masculino'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.Sexo = 'F' THEN 'Femenino'")
			loComandoSeleccionar.AppendLine(" 			END AS Sexo,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.tip_pag,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.nit,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.rif,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.movil,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.web,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.telefonos,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.fec_ini,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_pos,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.fax,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.fec_fin,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.correo,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.contacto,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_tip,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_zon,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_pai,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_est,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_ciu,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_cla,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_con,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_per,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_for,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_ven,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.cod_tra,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.matriz,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.dir_fis,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.dir_exa,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.dir_otr,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.dir_ent,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.actividad,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.are_neg,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.fec_nac,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.comentario,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.hor_caj,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.fiscal,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.nacional,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.mar_rec,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.crm_pos,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.pol_com,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.mon_sal,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.mon_cre,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.dia_cre,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.pro_pag,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_des,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_ali,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_int,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_pat,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_isl,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_ret,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_pat,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.por_otr,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.kilometros,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.generico,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Formal' THEN 'Formal'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Ordinario' THEN 'Ordinario'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Especial' THEN 'Especial'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Ocasional' THEN 'Ocasional'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Otro_1' THEN 'Otro 1'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Otro_2' THEN 'Otro 2'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Otro_3' THEN 'Otro 3'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Otro_4' THEN 'Otro 4'")
			loComandoSeleccionar.AppendLine(" 				WHEN Proveedores.tip_con = 'Otro_5' THEN 'Otro 5'")
			loComandoSeleccionar.AppendLine(" 			END AS tip_con,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.tipo,")
			loComandoSeleccionar.AppendLine(" 			Proveedores.clase,")
			loComandoSeleccionar.AppendLine(" 			Tipos_Proveedores.Nom_Tip,")
			loComandoSeleccionar.AppendLine(" 			Zonas.Nom_Zon,")
			loComandoSeleccionar.AppendLine(" 			Paises.Nom_Pai,")
			loComandoSeleccionar.AppendLine(" 			Estados.Nom_Est,")
			loComandoSeleccionar.AppendLine(" 			Ciudades.Nom_Ciu,")
			loComandoSeleccionar.AppendLine(" 			Clases_Proveedores.Nom_Cla,")
			loComandoSeleccionar.AppendLine(" 			Personas.Nom_Per,")
			loComandoSeleccionar.AppendLine(" 			Formas_Pagos.Nom_For,")
			loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven,")
			loComandoSeleccionar.AppendLine(" 			Transportes.Nom_Tra,")
			loComandoSeleccionar.AppendLine(" 			Conceptos.Nom_Con")
		
			
			loComandoSeleccionar.AppendLine(" FROM Proveedores")
			loComandoSeleccionar.AppendLine(" JOIN Tipos_Proveedores ON Tipos_Proveedores.Cod_Tip = Proveedores.Cod_Tip")
			loComandoSeleccionar.AppendLine(" JOIN Zonas ON Zonas.Cod_Zon = Proveedores.Cod_Zon")
			loComandoSeleccionar.AppendLine(" JOIN Paises ON Paises.Cod_Pai = Proveedores.Cod_Pai")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Estados ON Estados.Cod_Est = Proveedores.Cod_Est")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Ciudades ON Ciudades.Cod_Ciu = Proveedores.Cod_Ciu")
			loComandoSeleccionar.AppendLine(" JOIN Clases_Proveedores ON Clases_Proveedores.Cod_Cla = Proveedores.Cod_Cla")
			loComandoSeleccionar.AppendLine(" JOIN Personas ON Personas.Cod_Per = Proveedores.Cod_Per")
			loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Proveedores.Cod_For")
			loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Proveedores.Cod_Ven")
			loComandoSeleccionar.AppendLine(" JOIN Transportes ON Transportes.Cod_Tra = Proveedores.Cod_Tra")
			loComandoSeleccionar.AppendLine(" JOIN Conceptos ON Conceptos.Cod_Con = Proveedores.Cod_Con")
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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosProveedores", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosProveedores.ReportSource = loObjetoReporte

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
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'