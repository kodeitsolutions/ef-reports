'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_SID"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_SID
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	    Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("            Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("            Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("            Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("            Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("            Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("            Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("            Facturas.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("            Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("            Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("            Facturas.Por_Des1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("            Facturas.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Rec1       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("            Facturas.Cod_For, ")
            loComandoSeleccionar.AppendLine("            Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("            Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("            Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("            Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("            RTRIM(Articulos.Nom_Art) + '  ' ")
            loComandoSeleccionar.AppendLine("            + RTRIM(Renglones_Facturas.Comentario)    As  Nom_Art, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Mon_Imp1 As  Impuesto ,")
            loComandoSeleccionar.AppendLine("            CAST(Facturas.dis_imp AS XML) AS Impuestos")
            loComandoSeleccionar.AppendLine("INTO        #tmpFacturas")
            loComandoSeleccionar.AppendLine("FROM        Facturas ")
            loComandoSeleccionar.AppendLine("    JOIN    Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli ")
            loComandoSeleccionar.AppendLine("    JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = Facturas.Cod_For")
            loComandoSeleccionar.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven ")
            loComandoSeleccionar.AppendLine("    JOIN    Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  *, ")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/porcentaje)[1]',	'DECIMAL(28,10)'), 0) AS Impuesto_Por1,")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/base)[1]',	'DECIMAL(28,10)'), 0)   AS Impuesto_Bas1,")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/monto)[1]',	'DECIMAL(28,10)'), 0) AS Impuesto_Mon1,")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/porcentaje)[2]',	'DECIMAL(28,10)'), 0) AS Impuesto_Por2,")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/base)[2]',	'DECIMAL(28,10)'), 0)   AS Impuesto_Bas2,")
            loComandoSeleccionar.AppendLine("        ISNULL(Impuestos.value('(/impuestos/impuesto/monto)[2]',	'DECIMAL(28,10)'), 0) AS Impuesto_Mon2")
            loComandoSeleccionar.AppendLine("FROM    #tmpFacturas")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFacturas")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_SID", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Ventas_SID.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 04/06/13: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 04/09/13: Se ajustó el SELECT para que no trunque el nombre del artículo.            '
'-------------------------------------------------------------------------------------------'
