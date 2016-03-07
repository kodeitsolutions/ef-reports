Imports System.Data
Partial Class fOrden_Servicio_FSV3
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Substring(Renglones_Pedidos.Comentario,1,250)    As  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Pedidos.Documento  =   Renglones_Pedidos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art  =   Renglones_Pedidos.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrden_Servicio_FSV3", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrden_Servicio_FSV3.ReportSource = loObjetoReporte

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
' JFP: 21/11/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
