'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosVendedores"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosVendedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

				loComandoSeleccionar.AppendLine(" SELECT    ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Ven,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Adi,  ")
				loComandoSeleccionar.AppendLine(" 			CASE")
				loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'A' THEN 'Activo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'I' THEN 'Inactivo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'S' THEN 'Suspendido'")
				loComandoSeleccionar.AppendLine(" 			END AS Status,")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cedula,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Tip_Emp,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.ABC,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Nivel,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Prioridad,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Area,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Telefonos,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Directo,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Fax,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Movil,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Direccion,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Correo,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Web,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Fec_Ini,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Fec_Fin,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Fec_Nac,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Niv_Aca,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Tipo,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Clase,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Grupo,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Tip,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Pai,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Ciu,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Zon,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Ven,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Cob,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Imp,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Des,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Rec,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Ret,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Lic,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Mas,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Men,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Contable,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Foto,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Equ_Emp,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Comentario,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Notas,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Registro,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Cod_Suc,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Defecto,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Seleccion,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Otr1,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Otr2,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Otr3,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Otr4,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Otr5,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Sal,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Otr1,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Otr2,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Otr3,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Otr4,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Mon_Otr5,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Gan1,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Gan2,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Gan3,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Gan4,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Por_Gan5,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Usu_Cre,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Usu_Mod,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Reg_Mod,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Equ_Mod,  ")
				loComandoSeleccionar.AppendLine(" 			Vendedores.Equ_Cre,  ")
				loComandoSeleccionar.AppendLine(" 			Tipos_Vendedores.Nom_Tip,  ")
				loComandoSeleccionar.AppendLine(" 			Zonas.Nom_Zon  ")
				loComandoSeleccionar.AppendLine(" FROM Vendedores  ")
				loComandoSeleccionar.AppendLine(" JOIN Tipos_Vendedores ON Tipos_Vendedores.Cod_Tip = Vendedores.Cod_Tip")
				loComandoSeleccionar.AppendLine(" JOIN Zonas ON Zonas.Cod_Zon = Vendedores.Cod_Zon")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosVendedores", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosVendedores.ReportSource = loObjetoReporte

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
' CMS: 27/04/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'