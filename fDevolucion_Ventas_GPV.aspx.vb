Imports System.Data
Partial Class fDevolucion_Ventas_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT	    devoluciones_clientes.Cod_Cli,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos,  ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Nom_Cli        As  Nom_Gen,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Rif            As  Rif_Gen,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Nit            As  Nit_Gen,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Dir_Fis        As  Dir_Gen,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Telefonos      As  Tel_Gen,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Documento,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Fec_Fin,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Imp1,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Por_Des1,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Des1       As  Mon_Des,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Por_Rec1,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Mon_Rec1       AS  Mon_Rec,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Cod_For,  ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Cod_Ven,  ")
            loComandoSeleccionar.AppendLine("           devoluciones_clientes.Comentario,  ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Cod_Art,  ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Substring(renglones_dclientes.Comentario,1,450)    As  Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Renglon,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Can_Art1,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Cod_Uni,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Precio1,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Mon_Net  As  Neto,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Por_Imp1 As  Por_Imp,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Cod_Imp,  ")
            loComandoSeleccionar.AppendLine("           renglones_dclientes.Mon_Imp1 As  Impuesto  ")
            loComandoSeleccionar.AppendLine("FROM       devoluciones_clientes ")
            loComandoSeleccionar.AppendLine("	JOIN    renglones_dclientes ON renglones_dclientes.Documento = devoluciones_clientes.Documento ")
            loComandoSeleccionar.AppendLine("	JOIN    Articulos  ON Articulos.Cod_Art   =   renglones_dclientes.Cod_Art  ")
            loComandoSeleccionar.AppendLine("   JOIN    Clientes ON Clientes.Cod_Cli = devoluciones_clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("	JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = devoluciones_clientes.Cod_For ")
            loComandoSeleccionar.AppendLine("	JOIN    Vendedores ON Vendedores.Cod_Ven =  devoluciones_clientes.Cod_Ven  ")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDevolucion_Ventas_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDevolucion_Ventas_GPV.ReportSource = loObjetoReporte

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
' EAG: 15/08/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
