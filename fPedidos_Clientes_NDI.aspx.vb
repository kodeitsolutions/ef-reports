'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPedidos_Clientes_NDI"
'-------------------------------------------------------------------------------------------'
Partial Class fPedidos_Clientes_NDI

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Pedidos.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Clientes.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Clientes.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Pedidos ON (Pedidos.Documento  =   Renglones_Pedidos.Documento)")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON (Pedidos.Cod_Cli   =   Clientes.Cod_Cli) ")
            loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON (Pedidos.Cod_For =   Formas_Pagos.Cod_For) ")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON (Pedidos.Cod_Ven	=   Vendedores.Cod_Ven) ")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON (Articulos.Cod_Art	=   Renglones_Pedidos.Cod_Art)")
            loComandoSeleccionar.AppendLine(" WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPedidos_Clientes_NDI", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPedidos_Clientes_NDI.ReportSource = loObjetoReporte

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
' MAT: 03/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
