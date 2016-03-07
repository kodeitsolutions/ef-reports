Imports System.Data
Partial Class fFacturas_Ventas_Tire
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Facturas.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento  =   Renglones_Facturas.Documento AND ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Facturas.Status    <>   'Anulado' AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_Tire", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Ventas_Tire.ReportSource = loObjetoReporte

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
' GMO: 16/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GMO: 14/11/08: Adecuacion del Formato a Tire Solutions
'-------------------------------------------------------------------------------------------'
