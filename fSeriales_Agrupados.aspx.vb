'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeriales_Agrupados"
'-------------------------------------------------------------------------------------------'
Partial Class fSeriales_Agrupados
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Facturas.Cod_Cli							AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0) ")
            loComandoSeleccionar.AppendLine("				THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					(CASE WHEN (Facturas.Nom_Cli = '') ")
            loComandoSeleccionar.AppendLine("						THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("						ELSE Facturas.Nom_Cli ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("			END)										AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0) ")
            loComandoSeleccionar.AppendLine("				THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					(CASE WHEN (Facturas.Rif = '') ")
            loComandoSeleccionar.AppendLine("						THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("						ELSE Facturas.Rif ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("			END)										AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0) ")
            loComandoSeleccionar.AppendLine("				THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					(CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') ")
            loComandoSeleccionar.AppendLine("						THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("						ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("			END)										AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0) ")
            loComandoSeleccionar.AppendLine("				THEN Clientes.Telefonos ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					(CASE WHEN (Facturas.Telefonos = '') ")
            loComandoSeleccionar.AppendLine("						THEN Clientes.Telefonos ")
            loComandoSeleccionar.AppendLine("						ELSE Facturas.Telefonos ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("			END)										AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("			Clientes.Fax								AS Fax, ")
            loComandoSeleccionar.AppendLine("			Clientes.Generico							AS Generico, ")
            loComandoSeleccionar.AppendLine("			Facturas.Nom_Cli        					AS Nom_Gen, ")
            loComandoSeleccionar.AppendLine("			Facturas.Rif            					AS Rif_Gen, ")
            loComandoSeleccionar.AppendLine("			Facturas.Nit            					AS Nit_Gen, ")
            loComandoSeleccionar.AppendLine("			Facturas.Dir_Fis        					AS Dir_Gen, ")
            loComandoSeleccionar.AppendLine("			Facturas.Telefonos      					AS Tel_Gen, ")
            loComandoSeleccionar.AppendLine("			Facturas.Documento							AS Documento, ")
            loComandoSeleccionar.AppendLine("			Facturas.Fec_Ini							AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Facturas.Fec_Fin							AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Bru							AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Imp1							AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Imp1							AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Net							AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Des1							AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Des1							AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Rec1							AS Por_Rec1, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Rec1                       	AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("			Facturas.Cod_For                        	AS Cod_For, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Formas_Pagos.Nom_For,1,25)    	AS Nom_For, ")
            loComandoSeleccionar.AppendLine("			Facturas.Cod_Ven							AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Facturas.Comentario							AS Comentario,")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven							AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Cod_Art,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Generico = 0			  ")
            loComandoSeleccionar.AppendLine("				THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Notas ")
            loComandoSeleccionar.AppendLine("			END)										AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Renglon					AS Renglon, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Facturas.Can_Art1")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Can_Art2 ")
            loComandoSeleccionar.AppendLine("			END)										AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Facturas.Cod_Uni")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Cod_Uni2 ")
            loComandoSeleccionar.AppendLine("			END)										AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Facturas.Precio1")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2 ")
            loComandoSeleccionar.AppendLine("			END)										AS Precio1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Mon_Net					AS Neto,")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Por_Imp1					AS Por_Imp,")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Cod_Imp					AS Cod_Imp,")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Mon_Imp1					AS Impuesto,")
            loComandoSeleccionar.AppendLine("			Seriales.Origen								AS Origen,")
            loComandoSeleccionar.AppendLine("			Seriales.Renglon 							AS Renglon_Serial,")
            loComandoSeleccionar.AppendLine("			Seriales.Cod_Art 							AS Cod_Art_Serial,  ")
            loComandoSeleccionar.AppendLine("			Seriales.Nom_Art 							AS Nom_Art_Serial,")
            loComandoSeleccionar.AppendLine("			Seriales.Serial								AS Serial,")
            loComandoSeleccionar.AppendLine("			Seriales.Alm_Ent							AS Alm_Ent,")
            loComandoSeleccionar.AppendLine("			Seriales.Tip_Ent							AS Tip_Ent,")
            loComandoSeleccionar.AppendLine("			Seriales.Doc_Ent							AS Doc_Ent,")
            loComandoSeleccionar.AppendLine("			Seriales.Ren_Ent							AS Ren_Ent,")
            loComandoSeleccionar.AppendLine("			Seriales.Alm_Sal							AS Alm_Sal,")
            loComandoSeleccionar.AppendLine("			Seriales.Tip_Sal							AS Tip_Sal,")
            loComandoSeleccionar.AppendLine("			Seriales.Doc_Sal							AS Doc_Sal,")
            loComandoSeleccionar.AppendLine("			Seriales.Ren_Sal							AS Ren_Sal,")
            loComandoSeleccionar.AppendLine("			Seriales.Garantia							AS Garantia,")
            loComandoSeleccionar.AppendLine("			Seriales.Fec_Ini 							AS Fec_Fin_Serial,")
            loComandoSeleccionar.AppendLine("			Seriales.Fec_Fin 							AS Fec_Ini_Serial,")
            loComandoSeleccionar.AppendLine("			Seriales.Disponible							AS Disponible,")
            loComandoSeleccionar.AppendLine("			Seriales.Comentario							AS Comentario_Serial	")
            loComandoSeleccionar.AppendLine("INTO		#tmpSeriales")
            loComandoSeleccionar.AppendLine("FROM		Facturas")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas	ON Facturas.Documento	= Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes			ON Facturas.Cod_Cli		= Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("	JOIN	Formas_Pagos		ON Facturas.Cod_For		= Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("	JOIN	Vendedores			ON Facturas.Cod_Ven		= Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos			ON Articulos.Cod_Art	= Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN	Seriales			")
            loComandoSeleccionar.AppendLine("		ON	seriales.doc_sal	= facturas.documento")
            loComandoSeleccionar.AppendLine("		AND	Seriales.Ren_Sal	= Renglones_Facturas.Renglon ")
            loComandoSeleccionar.AppendLine("		AND	Seriales.tip_sal	= 'Facturas'")
            loComandoSeleccionar.AppendLine("WHERE	" &  goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT tmpExterno.Cod_Art, ")
            loComandoSeleccionar.AppendLine("(	SELECT	RTRIM(tmpInterno.Serial) + ', ' ")
            loComandoSeleccionar.AppendLine("	FROM	#tmpSeriales AS tmpInterno")
            loComandoSeleccionar.AppendLine("	WHERE	tmpInterno.Cod_Art = tmpExterno.Cod_Art")
            loComandoSeleccionar.AppendLine("	ORDER BY tmpInterno.Serial FOR XML PATH('')")
            loComandoSeleccionar.AppendLine(")	AS Seriales")
            loComandoSeleccionar.AppendLine("INTO #tmpAgrupados")
            loComandoSeleccionar.AppendLine("FROM #tmpSeriales AS tmpExterno")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	LEFT(#tmpAgrupados.Seriales, LEN(#tmpAgrupados.Seriales)-1) AS Seriales, ")
            loComandoSeleccionar.AppendLine("		#tmpSeriales.*		")
            loComandoSeleccionar.AppendLine("FROm	#tmpSeriales ")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpAgrupados ON  #tmpAgrupados.Cod_Art = #tmpSeriales.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY Ren_Sal")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLe #tmpSeriales")
            loComandoSeleccionar.AppendLine("DROP TABLe #tmpAgrupados")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
      

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeriales_Agrupados", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSeriales_Agrupados.ReportSource = loObjetoReporte

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
' CMS: 27/02/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 06/05/11: Agrupación de los seriales de un mismo artículo en un solo campo.			'
'-------------------------------------------------------------------------------------------'
' RJG: 16/05/11: Corrección: los seriales de un mismo artículo ocacionalmente aparecían		'
'				 duplicados (el grupo aparecía dos veces).									'
'-------------------------------------------------------------------------------------------'
