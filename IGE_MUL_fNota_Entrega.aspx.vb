Imports System.Data
Partial Class IGE_MUL_fNota_Entrega
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Entregas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Pos, ")
            loComandoSeleccionar.AppendLine("           Entregas.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Documento, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Bas1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Rec1       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Entregas.Comentario, ")
            loComandoSeleccionar.AppendLine("       CAST('' as char(400))           As  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Art, ")

            'loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            'loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Entregas.Notas) END AS Nom_Art,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Entregas.Comentario as Varchar(100)),3) <> '***' THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE LTRIM(SUBSTRING(CAST(Renglones_Entregas.Comentario as Varchar(100)),4,96)) END AS Nom_Art,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Entregas.Comentario as Varchar(100)),3) <> '***' THEN Renglones_Entregas.Comentario ")
            loComandoSeleccionar.AppendLine("			    ELSE '' END AS Comentario_Renglon,  ")


            'loComandoSeleccionar.AppendLine("           Renglones_Entregas.Comentario as Comentario_Renglon, ")




            'loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN (Articulos.Nom_Art + Substring(Renglones_Entregas.Comentario,1,500))")
            'loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Entregas.Notas + Substring(Renglones_Entregas.Comentario,1,500)) END AS Nom_Art,  ")
            ' loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Substring(Renglones_Entregas.Comentario,1,500)    As  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Entregas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Entregas.Documento  =   Renglones_Entregas.Documento AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Entregas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("IGE_MUL_fNota_Entrega", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvIGE_MUL_fNota_Entrega.ReportSource = loObjetoReporte

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
