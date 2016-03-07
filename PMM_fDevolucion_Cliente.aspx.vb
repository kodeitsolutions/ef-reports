Partial Class PMM_fDevolucion_Cliente
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Devoluciones_Clientes.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Devoluciones_Clientes.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Devoluciones_Clientes.Rif = '') THEN Clientes.Rif ELSE Devoluciones_Clientes.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Devoluciones_Clientes.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Devoluciones_Clientes.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Devoluciones_Clientes.Telefonos = '') THEN Clientes.Telefonos ELSE Devoluciones_Clientes.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Generico, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Documento, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Mon_Rec1                       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_For                        AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_dClientes.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Renglon, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_dClientes.Cod_Uni2='') THEN Renglones_dClientes.Can_Art1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_dClientes.Can_Art2 END) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_dClientes.Cod_Uni2='') THEN Renglones_dClientes.Cod_Uni")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_dClientes.Cod_Uni2 END) AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_dClientes.Cod_Uni2='') THEN Renglones_dClientes.Precio1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_dClientes.Precio1*Renglones_dClientes.Can_Uni2 END) AS Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Mon_Imp1 As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Tip_Ori As  TipoOrigen, ")
            loComandoSeleccionar.AppendLine("           Renglones_dClientes.Doc_Ori As  FacturaOrigen, ")
            loComandoSeleccionar.AppendLine("  			Seriales.Origen,")
            loComandoSeleccionar.AppendLine("  			Seriales.Renglon AS Renglon_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Cod_Art AS Cod_Art_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Nom_Art AS Nom_Art_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Alm_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Ren_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Alm_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Ren_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Garantia,")
            loComandoSeleccionar.AppendLine("  			Seriales.Fec_Ini AS Fec_Fin_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Fec_Fin AS Fec_Ini_Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Disponible,")
            loComandoSeleccionar.AppendLine("  			Operaciones_Lotes.Cod_Lot AS Lote,	")
            loComandoSeleccionar.AppendLine("  			Seriales.Comentario AS Comentario_Serial	")

            loComandoSeleccionar.AppendLine("INTO		#tmpDocumento		 ")

            loComandoSeleccionar.AppendLine("FROM       Devoluciones_Clientes ")

            loComandoSeleccionar.AppendLine("     JOIN Renglones_dClientes ON Devoluciones_Clientes.Documento = Renglones_dClientes.Documento")
            loComandoSeleccionar.AppendLine("     JOIN Clientes ON Devoluciones_Clientes.Cod_Cli     =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("     LEFT JOIN Seriales	ON	Seriales.Doc_Ent	=	Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine(" 				AND Seriales.Tip_Ent    =	'Devoluciones_Clientes'")
            loComandoSeleccionar.AppendLine(" 				AND Seriales.Ren_Ent	=   Renglones_dClientes.Renglon")
            loComandoSeleccionar.AppendLine("				AND Seriales.Cod_Art    =   Renglones_dClientes.Cod_Art")
            loComandoSeleccionar.AppendLine("     LEFT JOIN Formas_Pagos ON Devoluciones_Clientes.Cod_For   =   Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("     LEFT JOIN Vendedores ON Devoluciones_Clientes.Cod_Ven		=   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("     LEFT JOIN Articulos ON Articulos.Cod_Art  =  Renglones_dClientes.Cod_Art")
            loComandoSeleccionar.AppendLine("     LEFT JOIN Operaciones_Lotes ")
            loComandoSeleccionar.AppendLine("               ON Operaciones_Lotes.Num_doc   =  Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine("               AND Operaciones_Lotes.Tip_Doc = 'Devoluciones_Clientes' ")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")


            'loComandoSeleccionar.AppendLine(" WHERE     Devoluciones_Clientes.Documento      =   Renglones_dClientes.Documento ")
            'loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Cli    =   Clientes.Cod_Cli ")
            'loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_For    =   Formas_Pagos.Cod_For ")
            'loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Ven    =   Vendedores.Cod_Ven ")
            'loComandoSeleccionar.AppendLine("           AND Seriales.tip_sal='Facturas' and seriales.doc_sal=Devoluciones_Clientes.documento ")
            'loComandoSeleccionar.AppendLine("           AND Seriales.Ren_Sal =   Renglones_dClientes.Renglon  ")
            'loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Documento =   Operaciones_Lotes.Num_Doc and Operaciones_Lotes.Tip_Doc = 'Facturas'  ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art   =   Renglones_dClientes.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            loComandoSeleccionar.AppendLine("DECLARE @lcDocumento CHAR(10)")
            loComandoSeleccionar.AppendLine("SET @lcDocumento =  (SELECT TOP 1 Documento FROM #tmpDocumento)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER (ORDER BY Propiedades.Cod_Pro ASC) AS Contador_Propiedad,")
            loComandoSeleccionar.AppendLine("			Propiedades.Cod_Pro							AS Codigo_Propiedad,")
            loComandoSeleccionar.AppendLine("			Propiedades.Nom_Pro							AS Nombre_Propiedad,")
            loComandoSeleccionar.AppendLine("			Campos_Propiedades.Cod_Reg					AS Cod_Reg,")
            loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Car					AS Valor_Propiedad")
            loComandoSeleccionar.AppendLine("INTO		#tmpPropiedades")
            loComandoSeleccionar.AppendLine("FROM		Propiedades ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Campos_Propiedades ")
            loComandoSeleccionar.AppendLine("		ON	Campos_Propiedades.Cod_Pro = Propiedades.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Cod_Reg = @lcDocumento")
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Origen = 'devoluciones_clientes'")
            loComandoSeleccionar.AppendLine("WHERE		Propiedades.Status = 'A' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Modulo = 'Ventas' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Seccion = 'Operaciones' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Opcion = 'DevolucionesVentas'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpDocumento.*,")
            loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'SERMOT_VEN') AS SerialMotor,")
            loComandoSeleccionar.AppendLine("			COALESCE(   (SELECT TOP 1 FacturaOrigen ")
            loComandoSeleccionar.AppendLine("			            FROM #tmpDocumento")
            loComandoSeleccionar.AppendLine("			            WHERE TipoOrigen = 'facturas'")
            loComandoSeleccionar.AppendLine("			            ORDER BY Renglon ASC), '') AS FacturaOrigen")
            loComandoSeleccionar.AppendLine("FROM		#tmpDocumento")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PMM_fDevolucion_Cliente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvPMM_fDevolucion_Cliente.ReportSource = loObjetoReporte

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
' RJG: 27/08/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
