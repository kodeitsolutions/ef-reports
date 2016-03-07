Imports System.Data
Partial Class IGE_MIT_fFacturas_Ventas
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
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Pos, ")
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
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Bas1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Rec1       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("       CAST('' as char(400))           As  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")

            'loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            'loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Facturas.Notas) END AS Nom_Art,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Facturas.Comentario as Varchar(100)),3) <> '***' THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE LTRIM(SUBSTRING(CAST(Renglones_Facturas.Comentario as Varchar(100)),4,96)) END AS Nom_Art,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Facturas.Comentario as Varchar(100)),3) <> '***' THEN Renglones_Facturas.Comentario ")
            loComandoSeleccionar.AppendLine("			    ELSE '' END AS Comentario_Renglon,  ")


            'loComandoSeleccionar.AppendLine("           Renglones_Facturas.Comentario as Comentario_Renglon, ")




            'loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN (Articulos.Nom_Art + Substring(Renglones_Facturas.Comentario,1,500))")
            'loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Facturas.Notas + Substring(Renglones_Facturas.Comentario,1,500)) END AS Nom_Art,  ")
            ' loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Substring(Renglones_Facturas.Comentario,1,500)    As  Nom_Art, ")
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
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("IGE_MIT_fFacturas_Ventas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvIGE_MIT_fFacturas_Ventas.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Ajustes al select
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Adición de Descuentos y Recargos al Formato
'-------------------------------------------------------------------------------------------'
