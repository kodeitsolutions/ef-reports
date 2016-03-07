'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeriales_NRecepcionesProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fSeriales_NRecepcionesProveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

             loComandoSeleccionar.AppendLine(" SELECT    ")
            loComandoSeleccionar.AppendLine("			Recepciones.Documento,")
            loComandoSeleccionar.AppendLine("			Recepciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif,")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit,")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis,")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax,")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For,")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("  			Seriales.Cod_Art AS Cod_Art_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Nom_Art AS Nom_Art_Serial,")
            loComandoSeleccionar.AppendLine("           Seriales.Renglon AS Renglon_Serial, ")
            loComandoSeleccionar.AppendLine("  			Seriales.Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Ent")
            loComandoSeleccionar.AppendLine(" FROM      Recepciones")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Recepciones ON Renglones_Recepciones.Documento	=	Recepciones.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Recepciones.Cod_Pro     =   Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" JOIN Seriales	ON	Seriales.Doc_Ent	=	Recepciones.Documento")
            loComandoSeleccionar.AppendLine(" 				AND	 Seriales.tip_ent	=	'Recepciones'")
            loComandoSeleccionar.AppendLine(" 				AND	 Seriales.Ren_Ent	=   Renglones_Recepciones.Renglon")
            loComandoSeleccionar.AppendLine("				AND	Seriales.Cod_Art   =   Renglones_Recepciones.Cod_Art")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Formas_Pagos ON Recepciones.Cod_For   =   Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores ON Recepciones.Cod_Ven		=   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" WHERE		"  & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ORDER BY Renglones_Recepciones.Renglon,Seriales.Renglon ASC")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeriales_NRecepcionesProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSeriales_NRecepcionesProveedores.ReportSource = loObjetoReporte

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
' CMS: 04/03/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 23/03/11: Ajuste del Select (Eliminación de campos No Usados)
'-------------------------------------------------------------------------------------------'