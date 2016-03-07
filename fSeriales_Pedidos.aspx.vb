'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeriales_Pedidos"
'-------------------------------------------------------------------------------------------'
Partial Class fSeriales_Pedidos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT    ")
            loComandoSeleccionar.AppendLine("			Pedidos.Documento,")
            loComandoSeleccionar.AppendLine("			Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("           Clientes.Rif,")
            loComandoSeleccionar.AppendLine("           Clientes.Nit,")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis,")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos,")
            loComandoSeleccionar.AppendLine("           Clientes.Fax,")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For,")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("  			Seriales.Cod_Art AS Cod_Art_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Nom_Art AS Nom_Art_Serial,")
            loComandoSeleccionar.AppendLine("           Seriales.Renglon AS Renglon_Serial, ")
            loComandoSeleccionar.AppendLine("  			Seriales.Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Sal")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Pedidos ON Renglones_Pedidos.Documento	=	Pedidos.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Pedidos.Cod_Cli     =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" JOIN Seriales	ON	Seriales.Doc_Sal	=	Pedidos.Documento")
            loComandoSeleccionar.AppendLine(" 				AND	 Seriales.Tip_Sal	=	'Pedidos'")
            loComandoSeleccionar.AppendLine("				AND	Seriales.Cod_Art   =   Renglones_Pedidos.Cod_Art")
             loComandoSeleccionar.AppendLine("				AND  Renglones_Pedidos.Renglon	=	Seriales.Ren_Sal")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Formas_Pagos ON Pedidos.Cod_For   =   Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores ON Pedidos.Cod_Ven		=   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" WHERE		"  & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ORDER BY Renglones_Pedidos.Renglon,Seriales.Renglon ASC")
			
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeriales_Pedidos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSeriales_Pedidos.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MAT: 24/03/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
