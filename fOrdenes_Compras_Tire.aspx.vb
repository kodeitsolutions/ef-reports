Imports System.Data
Partial Class fOrdenes_Compras_Tire
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nom_Pro         As Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Rif             As Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nit             As Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dir_Fis         As Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Telefonos       As Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Uni1              As Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cos_Ult1     As Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Imp      As Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Imp1     As Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Imp1     As Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Compras, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Compras.Documento   =   Renglones_OCompras.Documento AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Pro     =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For     =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven     =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Status      <>  'Anulado' AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art           =   Renglones_OCompras.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Compras_Tire", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Compras_Tire.ReportSource = loObjetoReporte

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
' GMO: 08/11/08: Programacion Inicial
'-------------------------------------------------------------------------------------------'
