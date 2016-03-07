Imports System.Data
Partial Class CSJ_fFacturas_Ventas_Presentacion
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Generico, ")
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
            loComandoSeleccionar.AppendLine("           Facturas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Rec1                       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For                        AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Can_Art1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Can_Art2 END) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Cod_Uni")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Cod_Uni2 END) AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Precio1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2 END) AS Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Imp1 As  Impuesto ")
            
            'loComandoSeleccionar.AppendLine("INTO		#tmpDocumento		 ")

            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento      =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            'loComandoSeleccionar.AppendLine("DECLARE @lcDocumento CHAR(10)")
            'loComandoSeleccionar.AppendLine("SET @lcDocumento =  (SELECT TOp 1 Documento FROM #tmpDocumento)")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER (ORDER BY Propiedades.Cod_Pro ASC) AS Contador_Propiedad,")
            'loComandoSeleccionar.AppendLine("			Propiedades.Cod_Pro							AS Codigo_Propiedad,")
            'loComandoSeleccionar.AppendLine("			Propiedades.Nom_Pro							AS Nombre_Propiedad,")
            'loComandoSeleccionar.AppendLine("			Campos_Propiedades.Cod_Reg					AS Cod_Reg,")
            'loComandoSeleccionar.AppendLine("			Campos_Propiedades.Val_Car					AS Valor_Propiedad")
            'loComandoSeleccionar.AppendLine("INTO		#tmpPropiedades")
            'loComandoSeleccionar.AppendLine("FROM		Propiedades ")
            'loComandoSeleccionar.AppendLine("	LEFT JOIN Campos_Propiedades ")
            'loComandoSeleccionar.AppendLine("		ON	Campos_Propiedades.Cod_Pro = Propiedades.Cod_Pro")
            'loComandoSeleccionar.AppendLine("		AND Cod_Reg = @lcDocumento")
            'loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Origen = 'Facturas'")
            'loComandoSeleccionar.AppendLine("WHERE		Propiedades.Status = 'A' ")
            'loComandoSeleccionar.AppendLine("		AND	Propiedades.Modulo = 'Ventas' ")
            'loComandoSeleccionar.AppendLine("		AND	Propiedades.Seccion = 'Operaciones' ")
            'loComandoSeleccionar.AppendLine("		AND	Propiedades.Opcion = 'FacturasVentas'")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("SELECT		#tmpDocumento.*,")
            'loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'SERMOT_VEN') AS SerialMotor")
            'loComandoSeleccionar.AppendLine("FROM		#tmpDocumento")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CSJ_fFacturas_Ventas_Presentacion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_fFacturas_Ventas_Presentacion.ReportSource = loObjetoReporte

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
