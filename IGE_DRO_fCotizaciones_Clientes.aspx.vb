Imports System.Data
Partial Class IGE_DRO_fCotizaciones_Clientes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Pos, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bas1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Rec1       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("       CAST('' as char(400))           As  Mon_Let, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")

            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Cod_Dep = '011' THEN 1 ")
            loComandoSeleccionar.AppendLine("			    ELSE 0 END AS Sicos,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Cotizaciones.Comentario as Varchar(100)),3) <> '***' THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE LTRIM(SUBSTRING(CAST(Renglones_Cotizaciones.Comentario as Varchar(100)),4,96)) END AS Nom_Art,  ")

            loComandoSeleccionar.AppendLine("           CASE WHEN LEFT(CAST(Renglones_Cotizaciones.Comentario as Varchar(100)),3) <> '***' THEN Renglones_Cotizaciones.Comentario ")
            loComandoSeleccionar.AppendLine("			    ELSE '' END AS Comentario_Renglon,  ")

            loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Cotizaciones.Cod_Uni = Renglones_Cotizaciones.Cod_Uni2) ")
            loComandoSeleccionar.AppendLine("		  THEN Renglones_Cotizaciones.Can_Art1 ")
            loComandoSeleccionar.AppendLine("		  ELSE Renglones_Cotizaciones.Can_Art1/Renglones_Cotizaciones.Can_Uni2 ")
            loComandoSeleccionar.AppendLine("		END)                                                                        AS Can_Art1,")
            loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Cotizaciones.Cod_Uni = Renglones_Cotizaciones.Cod_Uni2) ")
            loComandoSeleccionar.AppendLine("		  THEN Renglones_Cotizaciones.Cod_Uni ")
            loComandoSeleccionar.AppendLine("		  ELSE Renglones_Cotizaciones.Cod_Uni2 ")
            loComandoSeleccionar.AppendLine("		END)                                                                        AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Cotizaciones.Cod_Uni = Renglones_Cotizaciones.Cod_Uni2) ")
            loComandoSeleccionar.AppendLine("		  THEN Renglones_Cotizaciones.Precio1 ")
            loComandoSeleccionar.AppendLine("		  ELSE Renglones_Cotizaciones.Precio1*Renglones_Cotizaciones.Can_Uni2 ")
            loComandoSeleccionar.AppendLine("		END)                                                                        AS Precio1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("           Unidades.Nom_Uni, ")

            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Unidades, ")
            loComandoSeleccionar.AppendLine("           Articulos ")

            'loComandoSeleccionar.AppendLine("	JOIN	Unidades   ")
            'loComandoSeleccionar.AppendLine("		ON	Unidades.Cod_Uni = Renglones_Cotizaciones.Cod_Uni2 ")

            loComandoSeleccionar.AppendLine(" WHERE     Cotizaciones.Documento  =   Renglones_Cotizaciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Uni2    =   Unidades.Cod_Uni AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Cotizaciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lnMontoNumero As Decimal
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                lnMontoNumero = CDec(loFilas.Item("Mon_Net"))
                loFilas.Item("Mon_Let") = goServicios.mConvertirMontoLetras(lnMontoNumero)

            Next loFilas

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("IGE_DRO_fCotizaciones_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvIGE_DRO_fCotizaciones_Clientes.ReportSource = loObjetoReporte

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
