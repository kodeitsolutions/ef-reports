Imports System.Data
Partial Class CSJ_fOrdenes_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            ''loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Nom_Pro ELSE ")
            ''loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            ''loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Rif ELSE ")
            ''loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            ''loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            ''loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            ''loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Telefonos ELSE ")
            ''loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")

            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            'loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            ''loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nom_Pro         As Nom_Gen, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Rif             As Rif_Gen, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nit             As Nit_Gen, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dir_Fis         As Dir_Gen, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Telefonos       As Tel_Gen, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Documento, ")

            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")

            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Uni, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Imp1, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dis_Imp, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Net, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Comentario, ")
            'loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art, ")
            ''loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            'loComandoSeleccionar.AppendLine("		CASE")
            'loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            'loComandoSeleccionar.AppendLine("			ELSE Renglones_OCompras.Notas")
            'loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Renglon, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Des      As Por_Des1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Comentario   As Comentario_Renglon, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Net      As Neto, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Imp      As Cod_Imp, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Imp1     As Por_Imp, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Imp1     As Impuesto ")
            'loComandoSeleccionar.AppendLine(" FROM      Ordenes_Compras, ")
            'loComandoSeleccionar.AppendLine("           Renglones_OCompras, ")
            'loComandoSeleccionar.AppendLine("           Proveedores, ")
            'loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            'loComandoSeleccionar.AppendLine("           Articulos ")
            'loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Compras.Documento   =   Renglones_OCompras.Documento AND ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Pro     =   Proveedores.Cod_Pro AND ")
            'loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For     =   Formas_Pagos.Cod_For AND ")
            'loComandoSeleccionar.AppendLine("           Articulos.Cod_Art           =   Renglones_OCompras.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            loComandoSeleccionar.AppendLine("SELECT		Ordenes_Compras.Cod_Pro				AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")

            loComandoSeleccionar.AppendLine("			Proveedores.Fax						AS Fax, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Nom_Pro         	AS Nom_Gen, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Rif             	AS Rif_Gen, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Nit             	AS Nit_Gen, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Dir_Fis         	AS Dir_Gen, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Telefonos       	AS Tel_Gen, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Documento			AS Documento, ")

            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Por_Des1 			AS Por_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Mon_Des1 			AS Mon_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Por_Rec1 			AS Por_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Mon_Rec1 			AS Mon_Rec1_Enc, ")

            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Cod_Uni			AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Fec_Ini				AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Fec_Fin				AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Mon_Bru				AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Por_Imp1			AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Dis_Imp				AS Dis_Imp, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Mon_Imp1			AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Mon_Net				AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Cod_For				AS Cod_For, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Comentario			AS Comentario, ")
            loComandoSeleccionar.AppendLine("           convert(varchar, Ordenes_compras.Fec_Ini, 110) 			As  Fecha_Texto, ")
            loComandoSeleccionar.AppendLine("			Formas_Pagos.Nom_For				AS Nom_For, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Compras.Cod_Ven				AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Cod_Art			AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("			CASE WHEN Articulos.Generico = 0 ")
            loComandoSeleccionar.AppendLine("				THEN Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_OCompras.Notas")
            loComandoSeleccionar.AppendLine("			END									AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Renglon			AS Renglon,			")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Can_Art1			AS Can_Art1,		")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Por_Des      	AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Precio1      	AS Precio1, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Comentario   	AS Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Mon_Net      	AS Neto, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Cod_Imp      	AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Por_Imp1     	AS Por_Imp, ")
            loComandoSeleccionar.AppendLine("			Renglones_OCompras.Mon_Imp1     	AS Impuesto")
            loComandoSeleccionar.AppendLine("INTO		#tmpDocumento		 ")
            loComandoSeleccionar.AppendLine("FROM		Ordenes_Compras		 ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_OCompras	ON Ordenes_Compras.Documento   =   Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores			ON Ordenes_Compras.Cod_Pro     =   Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN	Formas_Pagos		ON Ordenes_Compras.Cod_For     =   Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos			ON Articulos.Cod_Art           =   Renglones_OCompras.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE	" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lcDocumento CHAR(10)")
            loComandoSeleccionar.AppendLine("SET @lcDocumento =  (SELECT TOp 1 Documento FROM #tmpDocumento)")
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
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Origen = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("WHERE		Propiedades.Status = 'A' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Modulo = 'Compras' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Seccion = 'Operaciones' ")
            loComandoSeleccionar.AppendLine("		AND	Propiedades.Opcion = 'OrdenesCompra'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpDocumento.*,")
            loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'FIROCO1') AS Pedido_Por, ")
            loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'FIROCO2') AS Aprobado_Por")
            'loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'FIROCO3') AS Autorizado_Por")
            'loComandoSeleccionar.AppendLine("			(SELECT Valor_Propiedad FROM #tmpPropiedades WHERE Codigo_Propiedad = 'Val-REF') AS Valor_Referencia")
            loComandoSeleccionar.AppendLine("FROM		#tmpDocumento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpDocumento")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpPropiedades")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")





            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
                            lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                        End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CSJ_fOrdenes_Compras", laDatosReporte)
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_fOrdenes_Compras.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera datos de genericos de la Cotizacion cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 02/09/11: Adición de Comentario en Renglones
'-------------------------------------------------------------------------------------------'
