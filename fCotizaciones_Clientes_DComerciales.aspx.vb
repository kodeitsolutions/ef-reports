'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCotizaciones_Clientes_DComerciales"
'-------------------------------------------------------------------------------------------'
Partial Class fCotizaciones_Clientes_DComerciales

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
			
            loComandoSeleccionar.AppendLine("SELECT		Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END)		AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (CAST (Cotizaciones.Rif AS VARCHAR) = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END)					AS Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit																										AS Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (CAST (Cotizaciones.Dir_Fis AS VARCHAR) = '') THEN Clientes.Dir_Fis ELSE Cotizaciones.Dir_Fis END)		AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (CAST (Cotizaciones.Telefonos AS VARCHAR) = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END)	AS Telefonos, ")
            loComandoSeleccionar.AppendLine("			Clientes.Fax							AS Fax, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Nom_Cli                    AS Nom_Gen, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Rif                        AS Rif_Gen, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Nit                        AS Nit_Gen, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Dir_Fis                    AS Dir_Gen, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Telefonos                  AS Tel_Gen, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Documento					AS Documento, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Fec_Ini					AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Fec_Fin					AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Mon_Bru					AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Por_Des1					AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Por_Rec1					AS Por_Rec1, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Mon_Des1					AS Mon_Des1, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Mon_Rec1					AS Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Mon_Imp1					AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Mon_Net					AS Mon_Net	, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Cod_For					AS Cod_For	, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Dis_Imp					AS Dis_Imp	, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS Nom_For, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Cod_Ven					AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Comentario					AS Comentario, ")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven						AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Cod_Art			AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art						AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Renglon			AS Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Comentario		AS Comentario_renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Can_Art1			AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Cod_Uni			AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Precio1			AS Precio1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Por_Des			AS Por_Des, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Mon_Des			AS Mon_Des_Ren, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Mon_Net          AS Neto, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Por_Imp1         AS Por_Imp, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Cod_Imp			AS Cod_Imp,")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Mon_Imp1         AS Impuesto,")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Usa_Des_Com		AS Usa_Des_Com,")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Des_Com			AS Des_Com,")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Usa_Des_Vol		AS Usa_Des_Vol,")
            loComandoSeleccionar.AppendLine("			Renglones_Cotizaciones.Des_Vol			AS Des_Vol")
            loComandoSeleccionar.AppendLine("INTO		#tmpCotizacion ")
            loComandoSeleccionar.AppendLine("FROM		Cotizaciones ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cotizaciones ")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ")
            loComandoSeleccionar.AppendLine("		ON	Clientes.Cod_Cli = Cotizaciones.Cod_Cli")
            loComandoSeleccionar.AppendLine("	JOIN	Formas_Pagos ")
            loComandoSeleccionar.AppendLine("		ON	Formas_Pagos.Cod_For = Cotizaciones.Cod_For")
            loComandoSeleccionar.AppendLine("	JOIN	Vendedores ")
            loComandoSeleccionar.AppendLine("		ON	Vendedores.Cod_Ven = Cotizaciones.Cod_Ven")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos ")
            loComandoSeleccionar.AppendLine("		ON	Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	D.Renglon																										AS Renglon_Descuento,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/pre_ori').value('.','VARCHAR(30)')										AS Pre_Ori,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/mon_bas').value('.','VARCHAR(30)')										AS Mon_bas_Com,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/usa_pre').value('.','VARCHAR(30)')										AS Usa_Pre,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/bas_pre').value('.','VARCHAR(30)')										AS Bas_Pre,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/por_pre').value('.','VARCHAR(30)')										AS Por_Pre,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/mon_pre').value('.','VARCHAR(30)')										AS Mon_Pre,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('//descuentos/parametros/mon_des_com').value('.','VARCHAR(30)')									AS Mon_Des_Com,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[1]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com1,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[2]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com2,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[3]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com3,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[4]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com4,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[5]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com5,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[6]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com6,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[7]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com7,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[8]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com8,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[9]/por_des').value('.','VARCHAR(30)')									AS Por_Des_Com9,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[10]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com10,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[11]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com11,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[12]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com12,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[13]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com13,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[14]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com14,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[15]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com15,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[16]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com16,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[17]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com17,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[18]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com18,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[19]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com19,")
            loComandoSeleccionar.AppendLine("		D.Des_Com.query('(//descuentos/descuento)[20]/por_des').value('.','VARCHAR(30)')								AS Por_Des_Com20,")
            loComandoSeleccionar.AppendLine("		D.Des_Vol.query('//descuentos/parametros/mon_bas').value('.','VARCHAR(30)')										AS Mon_Bas_vol,")
            loComandoSeleccionar.AppendLine("		D.Des_Vol.query('//descuentos/parametros/mon_des_vol').value('.','VARCHAR(30)')									AS Mon_Des_vol,")
            loComandoSeleccionar.AppendLine("		D.Des_Vol.query('//descuentos/descuento/tip_des[text()=""Articulo""]/../por_des').value('.','VARCHAR(30)')		AS Por_Des_Art,")
            loComandoSeleccionar.AppendLine("		D.Des_Vol.query('//descuentos/descuento/tip_des[text()=""Departamento""]/../por_des').value('.','VARCHAR(30)')	AS Por_Des_Dep,")
            loComandoSeleccionar.AppendLine("		D.Des_Vol.query('//descuentos/descuento/tip_des[text()=""Segmento""]/../por_des').value('.','VARCHAR(30)')		AS Por_Des_Seg")
            loComandoSeleccionar.AppendLine("INTO	#tmpDescuentos		")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	Renglon, ")
            loComandoSeleccionar.AppendLine("					CAST(#tmpCotizacion.Des_Com AS XML) AS Des_Com, ")
            loComandoSeleccionar.AppendLine("					CAST(#tmpCotizacion.Des_Vol AS XML) AS Des_Vol")
            loComandoSeleccionar.AppendLine("			FROM	#tmpCotizacion")
            loComandoSeleccionar.AppendLine("		 ) AS D")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28,10)")
            loComandoSeleccionar.AppendLine("SET @lnCero = 0;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT C.*, ")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Pre_Ori) = ''		THEN @lnCero ELSE Pre_Ori END)			AS DECIMAL (28,10)) AS Pre_Ori,")
			loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Mon_bas_Com) = ''	THEN @lnCero ELSE Mon_bas_Com END)		AS DECIMAL (28,10)) AS Mon_Bas_Com,")
			loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Usa_Pre) = 'true'	THEN 1 ELSE 0 END) 						AS BIT)				AS Usa_Pre,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Bas_Pre) = '' 		THEN @lnCero ELSE Bas_Pre END) 			AS DECIMAL (28,10)) AS Bas_Pre,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Pre) = '' 		THEN @lnCero ELSE Por_Pre END) 			AS DECIMAL (28,10)) AS Por_Pre,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Mon_Pre) = '' 		THEN @lnCero ELSE Mon_Pre END) 			AS DECIMAL (28,10)) AS Mon_Pre,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Mon_Des_Com) = '' 	THEN @lnCero ELSE Mon_Des_Com END) 		AS DECIMAL (28,10)) AS Mon_Des_Com,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com1) = '' 	THEN @lnCero ELSE Por_Des_Com1 END) 	AS DECIMAL (28,10)) AS Por_Des_Com1,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com2) = '' 	THEN @lnCero ELSE Por_Des_Com2 END) 	AS DECIMAL (28,10)) AS Por_Des_Com2,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com3) = '' 	THEN @lnCero ELSE Por_Des_Com3 END) 	AS DECIMAL (28,10)) AS Por_Des_Com3,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com4) = '' 	THEN @lnCero ELSE Por_Des_Com4 END) 	AS DECIMAL (28,10)) AS Por_Des_Com4,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com5) = '' 	THEN @lnCero ELSE Por_Des_Com5 END) 	AS DECIMAL (28,10)) AS Por_Des_Com5,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com6) = '' 	THEN @lnCero ELSE Por_Des_Com6 END) 	AS DECIMAL (28,10)) AS Por_Des_Com6,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com7) = '' 	THEN @lnCero ELSE Por_Des_Com7 END) 	AS DECIMAL (28,10)) AS Por_Des_Com7,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com8) = '' 	THEN @lnCero ELSE Por_Des_Com8 END) 	AS DECIMAL (28,10)) AS Por_Des_Com8,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com9) = '' 	THEN @lnCero ELSE Por_Des_Com9 END) 	AS DECIMAL (28,10)) AS Por_Des_Com9,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com10) = '' 	THEN @lnCero ELSE Por_Des_Com10 END) 	AS DECIMAL (28,10)) AS Por_Des_Com10,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com11) = '' 	THEN @lnCero ELSE Por_Des_Com11 END) 	AS DECIMAL (28,10)) AS Por_Des_Com11,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com12) = '' 	THEN @lnCero ELSE Por_Des_Com12 END) 	AS DECIMAL (28,10)) AS Por_Des_Com12,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com13) = '' 	THEN @lnCero ELSE Por_Des_Com13 END) 	AS DECIMAL (28,10)) AS Por_Des_Com13,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com14) = '' 	THEN @lnCero ELSE Por_Des_Com14 END) 	AS DECIMAL (28,10)) AS Por_Des_Com14,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com15) = '' 	THEN @lnCero ELSE Por_Des_Com15 END) 	AS DECIMAL (28,10)) AS Por_Des_Com15,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com16) = '' 	THEN @lnCero ELSE Por_Des_Com16 END) 	AS DECIMAL (28,10)) AS Por_Des_Com16,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com17) = '' 	THEN @lnCero ELSE Por_Des_Com17 END) 	AS DECIMAL (28,10)) AS Por_Des_Com17,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com18) = '' 	THEN @lnCero ELSE Por_Des_Com18 END) 	AS DECIMAL (28,10)) AS Por_Des_Com18,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com19) = '' 	THEN @lnCero ELSE Por_Des_Com19 END) 	AS DECIMAL (28,10)) AS Por_Des_Com19,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Com20) = '' 	THEN @lnCero ELSE Por_Des_Com20 END) 	AS DECIMAL (28,10)) AS Por_Des_Com20,")	
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Mon_Bas_vol) = '' 	THEN @lnCero ELSE Mon_Bas_vol END) 		AS DECIMAL (28,10)) AS Mon_Bas_vol,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Mon_Des_vol) = '' 	THEN @lnCero ELSE Mon_Des_vol END) 		AS DECIMAL (28,10)) AS Mon_Des_vol,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Art) = '' 	THEN @lnCero ELSE Por_Des_Art END) 		AS DECIMAL (28,10)) AS Por_Des_Art,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Dep) = '' 	THEN @lnCero ELSE Por_Des_Dep END) 		AS DECIMAL (28,10)) AS Por_Des_Dep,")
            loComandoSeleccionar.AppendLine("		CAST( (CASE WHEN RTRIM(Por_Des_Seg) = '' 	THEN @lnCero ELSE Por_Des_Seg END) 		AS DECIMAL (28,10)) AS Por_Des_Seg")
            loComandoSeleccionar.AppendLine("FROM	#tmpCotizacion As C")
            loComandoSeleccionar.AppendLine("JOIN	#tmpDescuentos AS D ON D.Renglon_Descuento = C.Renglon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            
            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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
						If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
							lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
						End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")


			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_Clientes_DComerciales", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text38"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCotizaciones_Clientes_DComerciales.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 26/04/12: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
