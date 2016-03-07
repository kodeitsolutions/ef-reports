'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFacturas_OEnvio_OnyxTD"
'-------------------------------------------------------------------------------------------'
Partial Class rFacturas_OEnvio_OnyxTD
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            
            Dim loComandoSeleccionar As New StringBuilder()

   '         loComandoSeleccionar.AppendLine("SELECT	Documento ")
   '         loComandoSeleccionar.AppendLine("INTO	#tmpFactura")
   '         loComandoSeleccionar.AppendLine("FROM	Facturas")
			'loComandoSeleccionar.AppendLine("WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	" & lcParametro0Desde & " AS Documento ")
            loComandoSeleccionar.AppendLine("INTO	#tmpFactura")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpFactura.Documento,")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Nom_Cli)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Nom_Cli)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Nom_Cli)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Nom_Cli ")
            loComandoSeleccionar.AppendLine("			END AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Rif)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Rif)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Rif)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Rif ")
            loComandoSeleccionar.AppendLine("			END AS Rif, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Dir_Fis)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Dir_Fis)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Dir_Fis)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Dir_Fis ")
            loComandoSeleccionar.AppendLine("			END AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Telefonos)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Telefonos)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Telefonos)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Telefonos ")
            loComandoSeleccionar.AppendLine("			END AS Telefonos,			")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.For_Env)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.For_Env)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.For_Env)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.For_Env ")
            loComandoSeleccionar.AppendLine("			END AS For_Env, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Dir_Ent)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Dir_Ent)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Dir_Ent)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Dir_Ent ")
            loComandoSeleccionar.AppendLine("			END AS Dir_Ent,  ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Cod_Tra)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Cod_Tra)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Cod_Tra)")
            loComandoSeleccionar.AppendLine("				ELSE Facturas.Cod_Tra ")
            loComandoSeleccionar.AppendLine("			END AS Cod_Tra_Ori,  ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Comentario)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Comentario)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Comentario)")
            loComandoSeleccionar.AppendLine("				ELSE ' ' ")
            loComandoSeleccionar.AppendLine("			END AS Comentario_Origen  ")
            loComandoSeleccionar.AppendLine("INTO		#tmpOrigen")
            loComandoSeleccionar.AppendLine("FROM		Facturas")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpFactura ")
            loComandoSeleccionar.AppendLine("		ON	#tmpFactura.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ")
            loComandoSeleccionar.AppendLine("		ON	Facturas.Documento = Renglones_Facturas.Documento AND Renglones_Facturas.Renglon = 1")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ")
            loComandoSeleccionar.AppendLine("		ON	Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Cotizaciones ")
            loComandoSeleccionar.AppendLine("		ON (	Cotizaciones.Documento = Renglones_Facturas.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Tip_Ori = 'Cotizaciones')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Pedidos ")
            loComandoSeleccionar.AppendLine("		ON (	Pedidos.Documento = Renglones_Facturas.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Tip_Ori = 'Pedidos')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Entregas ")
            loComandoSeleccionar.AppendLine("		ON (	Entregas.Documento = Renglones_Facturas.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_Facturas.Tip_Ori = 'Entregas')")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Facturas.Cod_Cli												AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (#tmpOrigen.Nom_Cli = '') ")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("					ELSE #tmpOrigen.Nom_Cli ")
            loComandoSeleccionar.AppendLine("				END) END) 													AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (#tmpOrigen.Rif = '') ")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("					ELSE #tmpOrigen.Rif ")
            loComandoSeleccionar.AppendLine("            END) END)														AS Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit														AS Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) ")
            loComandoSeleccionar.AppendLine("				THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (SUBSTRING(#tmpOrigen.Dir_Fis,1, 200) = '') ")
            loComandoSeleccionar.AppendLine("							THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("							ELSE SUBSTRING(#tmpOrigen.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("						END) ")
            loComandoSeleccionar.AppendLine("				END)													AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (#tmpOrigen.Telefonos = '') ")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Telefonos ")
            loComandoSeleccionar.AppendLine("					ELSE #tmpOrigen.Telefonos ")
            loComandoSeleccionar.AppendLine("			END) END)														AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("			Clientes.Fax													AS Fax, ")
            loComandoSeleccionar.AppendLine("			Clientes.Generico												AS Generico, ")
            loComandoSeleccionar.AppendLine("			#tmpOrigen.Cod_Tra_Ori											AS Cod_Tra_Ori, ")
            loComandoSeleccionar.AppendLine("			Transportes.Nom_Tra												AS Nom_Tra_Ori, ")
            loComandoSeleccionar.AppendLine("			Transportes.Telefonos											AS Tel_Tra_Ori, ")
            loComandoSeleccionar.AppendLine("			Facturas.Documento												AS Documento, ")
            loComandoSeleccionar.AppendLine("			Facturas.Status 												AS Status, ")
            loComandoSeleccionar.AppendLine("			Facturas.Fec_Ini												AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Facturas.Fec_Fin												AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Facturas.Comentario												AS Comentario, ")
            loComandoSeleccionar.AppendLine("			#tmpOrigen.For_Env												AS For_Env,")
            loComandoSeleccionar.AppendLine("			#tmpOrigen.Dir_Ent												AS Dir_Ent, ")
            loComandoSeleccionar.AppendLine("           	Facturas.Notas													AS Notas, ")
            loComandoSeleccionar.AppendLine("           	Facturas.Cod_Tra												AS Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           	Transportes.Nom_Tra												AS Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           	Vendedores.Nom_Ven												AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           	Renglones_Facturas.Cod_Art										AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("           	CASE WHEN Articulos.Generico = 0 ")
            loComandoSeleccionar.AppendLine("           		THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Notas ")
            loComandoSeleccionar.AppendLine("			END 															AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           	Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("           		THEN Renglones_Facturas.Can_Art1")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Can_Art2 ")
            loComandoSeleccionar.AppendLine("			END) 															AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           	(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("           		THEN Renglones_Facturas.Cod_Uni")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Cod_Uni2 ")
            loComandoSeleccionar.AppendLine("			END)															AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           	Renglones_Facturas.Cod_Alm										AS Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           	RTRIM(LTRIM(Articulos.Garantia)) As Garantia,")
            loComandoSeleccionar.AppendLine("           	(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("           		THEN (Renglones_Facturas.Can_Art1 * Articulos.Peso) ")
            loComandoSeleccionar.AppendLine("				ELSE (Renglones_Facturas.Can_Art2 * Articulos.Peso) ")
            loComandoSeleccionar.AppendLine("			END)															AS Peso, ")
            loComandoSeleccionar.AppendLine("           	(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("           		THEN (Renglones_Facturas.Can_Art1 * Articulos.Volumen) ")
            loComandoSeleccionar.AppendLine("				ELSE (Renglones_Facturas.Can_Art2 * Articulos.Volumen) ")
            loComandoSeleccionar.AppendLine("			END)															AS Volumen, ")
            loComandoSeleccionar.AppendLine("           	Articulos.Cod_Ubi												AS Cod_Ubi, ")
            loComandoSeleccionar.AppendLine("           	Renglones_Facturas.Doc_Ori										AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("           	Renglones_Facturas.Tip_Ori										AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("			#tmpOrigen.Comentario_Origen									AS Comentario_Origen")
            loComandoSeleccionar.AppendLine("FROM		Facturas ")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpOrigen ON (#tmpOrigen.Documento	=   Facturas.Documento)")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON (Facturas.Documento	=   Renglones_Facturas.Documento)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Clientes ON (Facturas.Cod_Cli		=   Clientes.Cod_Cli)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Formas_Pagos ON (Facturas.Cod_For	=   Formas_Pagos.Cod_For)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Vendedores ON (Facturas.Cod_Ven		=   Vendedores.Cod_Ven)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Articulos ON (Articulos.Cod_Art		=   Renglones_Facturas.Cod_Art)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Transportes ON (Facturas.Cod_Tra	=   Transportes.Cod_Tra)")
            loComandoSeleccionar.AppendLine("ORDER BY Renglones_Facturas.Renglon ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFactura")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpOrigen")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
   

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcCadenaComentario As String = ""
            Dim lcComentario As String

            lcCadenaComentario = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcComentario = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Comentario_Origen")

                If lcComentario = "" Then
                    Continue For
                End If
				lcCadenaComentario = lcCadenaComentario.Trim & lcComentario & ",  "
				
            Next lnNumeroFila

            lcCadenaComentario = lcCadenaComentario & ")"
            lcCadenaComentario = lcCadenaComentario.Replace("(,", "(")
            lcCadenaComentario = lcCadenaComentario.Replace(".", ",")
            lcCadenaComentario = lcCadenaComentario.Replace(",)", ")")
            lcCadenaComentario = lcCadenaComentario.Replace("(,", "(")
            lcCadenaComentario = lcCadenaComentario.Replace(", ", "")
            lcCadenaComentario = lcCadenaComentario.Replace(",", ", ")

   			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_OEnvio_OnyxTD", laDatosReporte)

			CType(loObjetoReporte.ReportDefinition.ReportObjects("Text24"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcCadenaComentario.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_OEnvio_OnyxTD.ReportSource = loObjetoReporte

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
'-----------------------------------------------------------------------------------------------'
' Fin del codigo																				'
'-----------------------------------------------------------------------------------------------'
' RJG: 24/01/13: Codigo inicial, a partir de fFacturas_OEnvio_OnyxTD.							'
'-----------------------------------------------------------------------------------------------'
