'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_MARCAPI"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_MARCAPI
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcDocumento CHAR(10); ")
            loComandoSeleccionar.Append    ("SET @lcDocumento = (SELECT TOP 1 Documento FROM facturas WHERE ")
            loComandoSeleccionar.Append    (cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(");")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Obtiene las formas de pago:")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFormasPago(Tip_Ope CHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Num_Doc CHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Mon_Net DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("                            Cod_Tar CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Nom_Tar CHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Cod_Ban CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Nom_Ban CHAR(100) COLLATE DATABASE_DEFAULT);")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFormasPago(Tip_Ope, Num_Doc ,Mon_Net, Cod_Tar, Nom_Tar, Cod_Ban, Nom_Ban)")
            loComandoSeleccionar.AppendLine("SELECT      detalles_cobros.tip_ope, ")
            loComandoSeleccionar.AppendLine("            detalles_cobros.num_doc, ")
            loComandoSeleccionar.AppendLine("            detalles_cobros.mon_net, ")
            loComandoSeleccionar.AppendLine("            detalles_cobros.cod_tar,")
            loComandoSeleccionar.AppendLine("            COALESCE(Tarjetas.Nom_tar, detalles_cobros.cod_tar),")
            loComandoSeleccionar.AppendLine("            detalles_cobros.cod_ban,")
            loComandoSeleccionar.AppendLine("            COALESCE(bancos.Nom_Ban, detalles_cobros.cod_ban)")
            loComandoSeleccionar.AppendLine("FROM        Renglones_cobros ")
            loComandoSeleccionar.AppendLine("    JOIN    cobros ")
            loComandoSeleccionar.AppendLine("        ON  cobros.documento = Renglones_cobros.documento")
            loComandoSeleccionar.AppendLine("        AND cobros.status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("    JOIN detalles_cobros ")
            loComandoSeleccionar.AppendLine("        ON detalles_cobros.documento = cobros.documento")
            loComandoSeleccionar.AppendLine("    LEFT JOIN bancos ")
            loComandoSeleccionar.AppendLine("        ON bancos.cod_ban = detalles_cobros.cod_ban")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Tarjetas ")
            loComandoSeleccionar.AppendLine("        ON Tarjetas.Cod_Tar = detalles_cobros.Cod_Tar")
            loComandoSeleccionar.AppendLine("WHERE   Renglones_cobros.doc_ori = @lcDocumento")
            loComandoSeleccionar.AppendLine("    AND Renglones_cobros.cod_tip = 'FACT';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Obtiene la factura y sus datos")
            loComandoSeleccionar.AppendLine("SELECT	    Facturas.Cod_Cli                AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine("            Clientes.Nom_Cli                AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("            Clientes.Rif                    AS Rif, ")
            loComandoSeleccionar.AppendLine("            Clientes.Nit                    AS Nit, ")
            loComandoSeleccionar.AppendLine("            Clientes.Dir_Fis                AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("            Clientes.Telefonos              AS Telefonos, ")
            loComandoSeleccionar.AppendLine("            Clientes.Fax                    AS Fax, ")
            loComandoSeleccionar.AppendLine("            Facturas.Nom_Cli                AS Nom_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Rif                    AS Rif_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Nit                    AS Nit_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Dir_Fis                AS Dir_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Telefonos              AS Tel_Gen, ")
            loComandoSeleccionar.AppendLine("            Facturas.Documento              AS Documento, ")
            loComandoSeleccionar.AppendLine("            Facturas.Fec_Ini                AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("            Facturas.Fec_Fin                AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Bru                AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Imp1               AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Net                AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Net                AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("            Facturas.Por_Des1               AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Des1               AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("            Facturas.Por_Rec1               AS Por_Rec1, ")
            loComandoSeleccionar.AppendLine("            Facturas.Mon_Rec1               AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("            Facturas.Cod_For                AS Cod_For, ")
            loComandoSeleccionar.AppendLine("            Formas_Pagos.Nom_For            AS Nom_For, ")
            loComandoSeleccionar.AppendLine("            Facturas.Cod_Ven                AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("            Facturas.Comentario             AS Comentario, ")
            loComandoSeleccionar.AppendLine("            Vendedores.Nom_Ven              AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Art      AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("            Articulos.Nom_Art + SUBSTRING(Renglones_Facturas.Comentario,1,250)    AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Renglon      AS Renglon, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Can_Art1     AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Uni      AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Precio1      AS Precio1, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Mon_Net      AS Neto, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Por_Imp1     AS Por_Imp, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Cod_Imp      AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("            Renglones_Facturas.Mon_Imp1     AS Impuesto,")
            loComandoSeleccionar.AppendLine("            COALESCE(Efectivo.Mon_Efe, 0)   AS Efectivo,")
            loComandoSeleccionar.AppendLine("            COALESCE(Cheques.Num_Doc, '')   AS Cheque,")
            loComandoSeleccionar.AppendLine("            COALESCE(Cheques.Nom_Ban, '')   AS Banco,")
            loComandoSeleccionar.AppendLine("            COALESCE(Tarjetas.Nom_Tar, '')  AS Tarjeta,")
            loComandoSeleccionar.AppendLine("            COALESCE(Otros.Otro_Tipo, '')   AS Otro_Tipo")
            loComandoSeleccionar.AppendLine("FROM        Facturas")
            loComandoSeleccionar.AppendLine("    JOIN    Renglones_Facturas ON Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Clientes ON Facturas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("    JOIN    Formas_Pagos ON Facturas.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("    JOIN    Vendedores ON Facturas.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("    JOIN    Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT SUM(Mon_Net) AS Mon_Efe")
            loComandoSeleccionar.AppendLine("                FROM #tmpFormasPago")
            loComandoSeleccionar.AppendLine("                WHERE Tip_Ope= 'Efectivo'")
            loComandoSeleccionar.AppendLine("            ) AS Efectivo ON 1=1")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  TOP 1 Num_Doc, Nom_Ban ")
            loComandoSeleccionar.AppendLine("                FROM    #tmpFormasPago ")
            loComandoSeleccionar.AppendLine("                WHERE   Tip_Ope = 'Cheque'")
            loComandoSeleccionar.AppendLine("            ) AS Cheques ON 1=1")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  TOP 1 RTRIM(Nom_Tar) + ' ' + Num_Doc AS Nom_Tar ")
            loComandoSeleccionar.AppendLine("                FROM    #tmpFormasPago ")
            loComandoSeleccionar.AppendLine("                WHERE   Tip_Ope = 'Tarjeta'")
            loComandoSeleccionar.AppendLine("            ) AS Tarjetas ON 1=1")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  TOP 1 RTRIM(Tip_Ope) + ' ' + Num_Doc AS Otro_Tipo ")
            loComandoSeleccionar.AppendLine("                FROM    #tmpFormasPago ")
            loComandoSeleccionar.AppendLine("                WHERE   Tip_Ope NOT IN ('Efectivo', 'Cheque', 'Tarjeta')")
            loComandoSeleccionar.AppendLine("            ) AS Otros ON 1=1")
            loComandoSeleccionar.AppendLine("WHERE       Facturas.Documento = @lcDocumento;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFormasPago;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_MARCAPI", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Ventas_MARCAPI.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 17/05/13: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 25/06/13: Se agregaron las formas de pago del cobro asociado.                        '
'-------------------------------------------------------------------------------------------'
