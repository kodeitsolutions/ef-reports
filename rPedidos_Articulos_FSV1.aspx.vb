Imports System.Data
Partial Class rPedidos_Articulos_FSV1
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcComandoSelect As String

            lcComandoSelect = " SELECT  Articulos.Cod_Art, " _
            & " Articulos.Nom_art, " _
            & " Articulos.Cod_Mar, " _
            & " Articulos.Status, " _
            & " Pedidos.Documento, " _
            & " Pedidos.Cod_Cli, " _
            & " Clientes.Nom_Cli, " _
            & " Pedidos.Fec_Ini, " _
            & " Pedidos.Cod_Ven, " _
            & " Renglones_Pedidos.Renglon, " _
            & " Renglones_Pedidos.Cod_Alm, " _
            & " Renglones_Pedidos.Can_Pen1  As  Can_Art1, " _
            & " Renglones_Pedidos.Cod_Uni, " _
            & " Renglones_Pedidos.Precio1, " _
            & " Renglones_Pedidos.Por_Des, " _
            & " Renglones_Pedidos.Mon_Net " _
         & " From Articulos, " _
            & " Pedidos, " _
            & " Renglones_Pedidos, " _
            & " Clientes, " _
            & " Vendedores, " _
            & " Almacenes, " _
            & " Marcas " _
         & " WHERE  Articulos.Cod_Art           =   Renglones_Pedidos.Cod_Art " _
            & " And Renglones_Pedidos.Documento =   Pedidos.Documento " _
            & " And Articulos.Cod_Mar           =   Marcas.Cod_Mar " _
            & " And Pedidos.Cod_Cli             =   Clientes.Cod_Cli " _
            & " And Pedidos.Cod_Ven             =   Vendedores.Cod_Ven " _
            & " And Renglones_Pedidos.Cod_Alm   =   Almacenes.Cod_Alm " _
            & " And Renglones_Pedidos.Can_Pen1 <> 0 " _
            & " And Pedidos.Status <> 'Anulado' " _
            & " And (substring(Renglones_Pedidos.Cod_Art,1,3)='SOP' OR substring(Renglones_Pedidos.Cod_Art,1,3)='PRO' ) " _
          & " ORDER BY Pedidos.Cod_Cli, Articulos.Cod_Art"


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("rPedidos_Articulos_FSV1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Articulos_FSV1.ReportSource = loObjetoReporte

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
' JFP: 01/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
