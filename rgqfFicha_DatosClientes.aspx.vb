'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rgqfFicha_DatosClientes"
'-------------------------------------------------------------------------------------------'
Partial Class rgqfFicha_DatosClientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")				
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_cli,")
			loComandoSeleccionar.AppendLine(" 			Clientes.nom_cli,")
            loComandoSeleccionar.AppendLine(" 			Campos_Extras.caracter1,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine(" 			Clientes.tip_cli,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.crm_sec = 'P' THEN 'Producción'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.crm_sec = 'C' THEN 'Comercio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.crm_sec = 'S' THEN 'Servicio'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.crm_sec = 'G' THEN 'Gobierno'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.crm_sec = 'O' THEN 'Otro'")
			loComandoSeleccionar.AppendLine(" 			END AS crm_sec,")
			loComandoSeleccionar.AppendLine(" 			Clientes.abc,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Sexo = 'X' THEN 'No Aplica'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Sexo = 'M' THEN 'Masculino'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.Sexo = 'F' THEN 'Femenino'")
			loComandoSeleccionar.AppendLine(" 			END AS Sexo,")
			loComandoSeleccionar.AppendLine(" 			Clientes.tip_pag,")
			loComandoSeleccionar.AppendLine(" 			Clientes.nit,")
			loComandoSeleccionar.AppendLine(" 			Clientes.rif,")
			loComandoSeleccionar.AppendLine(" 			Clientes.movil,")
			loComandoSeleccionar.AppendLine(" 			Clientes.web,")
			loComandoSeleccionar.AppendLine(" 			Clientes.telefonos,")
			loComandoSeleccionar.AppendLine(" 			Clientes.fec_ini,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_pos,")
			loComandoSeleccionar.AppendLine(" 			Clientes.fax,")
			loComandoSeleccionar.AppendLine(" 			Clientes.fec_fin,")
			loComandoSeleccionar.AppendLine(" 			Clientes.correo,")
			loComandoSeleccionar.AppendLine(" 			Clientes.contacto,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_tip,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_zon,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_pai,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_est,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_ciu,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_cla,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_con,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_per,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_for,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_ven,")
			loComandoSeleccionar.AppendLine(" 			Clientes.cod_tra,")
			loComandoSeleccionar.AppendLine(" 			Clientes.matriz,")
			loComandoSeleccionar.AppendLine(" 			Clientes.dir_fis,")
			loComandoSeleccionar.AppendLine(" 			Clientes.dir_exa,")
			loComandoSeleccionar.AppendLine(" 			Clientes.dir_otr,")
			loComandoSeleccionar.AppendLine(" 			Clientes.dir_ent,")
			loComandoSeleccionar.AppendLine(" 			Clientes.actividad,")
			loComandoSeleccionar.AppendLine(" 			Clientes.are_neg,")
			loComandoSeleccionar.AppendLine(" 			Clientes.fec_nac,")
			loComandoSeleccionar.AppendLine(" 			Clientes.comentario,")
			loComandoSeleccionar.AppendLine(" 			Clientes.hor_caj,")
			loComandoSeleccionar.AppendLine(" 			Clientes.fiscal,")
			loComandoSeleccionar.AppendLine(" 			Clientes.nacional,")
			loComandoSeleccionar.AppendLine(" 			Clientes.mar_rec,")
			loComandoSeleccionar.AppendLine(" 			Clientes.crm_pos,")
			loComandoSeleccionar.AppendLine(" 			Clientes.pol_com,")
			loComandoSeleccionar.AppendLine(" 			Clientes.mon_sal,")
			loComandoSeleccionar.AppendLine(" 			Clientes.mon_cre,")
			loComandoSeleccionar.AppendLine(" 			Clientes.dia_cre,")
			loComandoSeleccionar.AppendLine(" 			Clientes.pro_pag,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_des,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_ali,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_int,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_pat,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_isl,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_ret,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_pat,")
			loComandoSeleccionar.AppendLine(" 			Clientes.por_otr,")
			loComandoSeleccionar.AppendLine(" 			Clientes.kilometros,")
			loComandoSeleccionar.AppendLine(" 			Clientes.generico,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Formal' THEN 'Formal'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Ordinario' THEN 'Ordinario'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Especial' THEN 'Especial'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Ocasional' THEN 'Ocasional'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Otro_1' THEN 'Otro 1'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Otro_2' THEN 'Otro 2'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Otro_3' THEN 'Otro 3'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Otro_4' THEN 'Otro 4'")
			loComandoSeleccionar.AppendLine(" 				WHEN Clientes.tip_con = 'Otro_5' THEN 'Otro 5'")
			loComandoSeleccionar.AppendLine(" 			END AS tip_con,")
			loComandoSeleccionar.AppendLine(" 			Clientes.tipo,")
			loComandoSeleccionar.AppendLine(" 			Clientes.clase,")
			loComandoSeleccionar.AppendLine(" 			Tipos_Clientes.Nom_Tip,")
			loComandoSeleccionar.AppendLine(" 			Zonas.Nom_Zon,")
			loComandoSeleccionar.AppendLine(" 			Paises.Nom_Pai,")
			loComandoSeleccionar.AppendLine(" 			Estados.Nom_Est,")
			loComandoSeleccionar.AppendLine(" 			Ciudades.Nom_Ciu,")
			loComandoSeleccionar.AppendLine(" 			Clases_Clientes.Nom_Cla,")
			loComandoSeleccionar.AppendLine(" 			Personas.Nom_Per,")
			loComandoSeleccionar.AppendLine(" 			Formas_Pagos.Nom_For,")
			loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven,")
			loComandoSeleccionar.AppendLine(" 			Transportes.Nom_Tra,")
			loComandoSeleccionar.AppendLine(" 			Conceptos.Nom_Con")
			loComandoSeleccionar.AppendLine(" FROM Clientes")
			loComandoSeleccionar.AppendLine(" JOIN Tipos_Clientes ON Tipos_Clientes.Cod_Tip = Clientes.Cod_Tip")
			loComandoSeleccionar.AppendLine(" JOIN Zonas ON Zonas.Cod_Zon = Clientes.Cod_Zon")
			loComandoSeleccionar.AppendLine(" JOIN Paises ON Paises.Cod_Pai = Clientes.Cod_Pai")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Estados ON Estados.Cod_Est = Clientes.Cod_Est")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Ciudades ON Ciudades.Cod_Ciu = Clientes.Cod_Ciu")
			loComandoSeleccionar.AppendLine(" JOIN Clases_Clientes ON Clases_Clientes.Cod_Cla = Clientes.Cod_Cla")
			loComandoSeleccionar.AppendLine(" JOIN Personas ON Personas.Cod_Per = Clientes.Cod_Per")
			loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Clientes.Cod_For")
			loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Clientes.Cod_Ven")
			loComandoSeleccionar.AppendLine(" JOIN Transportes ON Transportes.Cod_Tra = Clientes.Cod_Tra")
			loComandoSeleccionar.AppendLine(" JOIN Conceptos ON Conceptos.Cod_Con = Clientes.Cod_Con")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Campos_Extras ON Campos_Extras.cod_reg = Clientes.Cod_cli and Campos_Extras.origen = 'Clientes'")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("rgqfFicha_DatosClientes", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrgqfFicha_DatosClientes.ReportSource = loObjetoReporte

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
' RGQ: 13/10/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
