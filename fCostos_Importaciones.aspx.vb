'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCostos_Importaciones"
'-------------------------------------------------------------------------------------------'
Partial Class fCostos_Importaciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()
            
            Dim llIncluyeImpuesto AS Boolean = goOpciones.mObtener("INCIMPORDP","L")
	
			loConsulta.AppendLine("")
			loConsulta.AppendLine("CREATE TABLE #tmpDocumentoActual(   Documento CHAR(10) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                                    Origen CHAR(100) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                                    Adicional CHAR(250) COLLATE DATABASE_DEFAULT")
			loConsulta.AppendLine("                                    );")
			loConsulta.AppendLine("INSERT INTO #tmpDocumentoActual(Documento, Origen, Adicional)")
			loConsulta.AppendLine("SELECT		Documento, Origen, Adicional")
			loConsulta.AppendLine("FROM		Importaciones")
			loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loConsulta.AppendLine("")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("/* Tabla General.																			*/")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("CREATE TABLE #tmpImportacion(")
			loConsulta.AppendLine("    --Encabezado y Pie")
			loConsulta.AppendLine("    Tipo INT, ")
			loConsulta.AppendLine("    Origen CHAR(100), Documento CHAR(10), Adicional CHAR(250), Renglon INT, ")
			loConsulta.AppendLine("    Status CHAR(15), Referencia CHAR(30), Expediente CHAR(30), Fec_Ini DATETIME, Cod_Mon CHAR(10), Tasa DECIMAL(28,10),")
			loConsulta.AppendLine("    Cod_Pro CHAR(10), Nom_Pro VARCHAR(MAX), Comentario VARCHAR(MAX), Tip_Dis CHAR(30), Tip_Ara CHAR(30), Mon_Bru DECIMAL(28,10),")
			loConsulta.AppendLine("    Mon_Gas_Fij DECIMAL(28,10), Mon_Gas_Adi DECIMAL(28,10), Mon_Com_Pag DECIMAL(28,10), ")
			loConsulta.AppendLine("    Mon_Gas_Com DECIMAL(28,10), Mon_Ara DECIMAL(28,10), Mon_Net DECIMAL(28,10),")
			loConsulta.AppendLine("    Tas_Emi DECIMAL(28,10), Tas_Car DECIMAL(28,10), Tas_Bar DECIMAL(28,10), Tas_Otr1 DECIMAL(28,10),")
			loConsulta.AppendLine("    Tas_Otr2 DECIMAL(28,10), Tas_Otr3 DECIMAL(28,10), Tas_Otr4 DECIMAL(28,10), Tas_Otr5 DECIMAL(28,10),	")
			loConsulta.AppendLine("    --Renglones ")
			loConsulta.AppendLine("    Ren_Cod_Art CHAR(30), Ren_Nom_Art VARCHAR(MAX), Ren_Cod_Alm CHAR(10), ")
			loConsulta.AppendLine("    Ren_Can_Art1 DECIMAL(28,10), Ren_Precio1 DECIMAL(28,10), Ren_Mon_Bru DECIMAL(28,10),")
			loConsulta.AppendLine("    Ren_Mon_Seg DECIMAL(28,10), Ren_Mon_Fle DECIMAL(28,10), Ren_Mon_Alm DECIMAL(28,10),")
			loConsulta.AppendLine("    Ren_Mon_Ipt DECIMAL(28,10), Ren_Mon_Por DECIMAL(28,10), Ren_Mon_Tra DECIMAL(28,10),")
			loConsulta.AppendLine("    Ren_Mon_Per DECIMAL(28,10), Ren_Mon_Ban DECIMAL(28,10), Ren_Mon_Adu DECIMAL(28,10), Ren_Mon_Arc DECIMAL(28,10),")
			loConsulta.AppendLine("    Ren_Mon_Ots1 DECIMAL(28,10), Ren_Mon_Ots2 DECIMAL(28,10), Ren_Mon_Ots3 DECIMAL(28,10),")
			loConsulta.AppendLine("    --Detalles")
			loConsulta.AppendLine("    Det_Grupo CHAR(30), Det_Afe_Cos BIT, Det_Cod_Gas CHAR(10), Det_Nom_Gas VARCHAR(MAX), Det_Mon_Gas DECIMAL(28,10), ")
			loConsulta.AppendLine("    Det_Tip_Ori CHAR(30), Det_Doc_Ori CHAR(10), Det_Ren_Ori INT, ")
			loConsulta.AppendLine("    Det_Cod_Art CHAR(30), Det_Nom_Art VARCHAR(MAX), Det_Can_Art DECIMAL(28,10), ")
			loConsulta.AppendLine("    Det_Precio1 DECIMAL(28,10), Det_Cod_Con CHAR(10), Det_Nom_Con VARCHAR(MAX), ")
			loConsulta.AppendLine("    Det_Mon_Deb DECIMAL(28,10), Det_Mon_Hab DECIMAL(28,10), Det_Cod_Imp CHAR(10), ")
			loConsulta.AppendLine("    Det_Por_Imp DECIMAL(28,10), Det_Mon_Imp DECIMAL(28,10), Det_Mon_Bru DECIMAL(28,10), ")
			loConsulta.AppendLine("    Det_Mon_Net DECIMAL(28,10), Det_Tip_Gas CHAR(50), Det_Con_Gas CHAR(50)")
			loConsulta.AppendLine(");")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("/* Renglones.																				*/")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("INSERT INTO #tmpImportacion")
			loConsulta.AppendLine("	(	Tipo, Origen, Documento, Adicional, ")
			loConsulta.AppendLine("		Renglon, Ren_Cod_Art, Ren_Nom_Art, Ren_Cod_Alm, ")
			loConsulta.AppendLine("		Ren_Can_Art1, Ren_Precio1, Ren_Mon_Bru,")
			loConsulta.AppendLine("        Ren_Mon_Seg, Ren_Mon_Fle, Ren_Mon_Alm,")
			loConsulta.AppendLine("        Ren_Mon_Ipt, Ren_Mon_Por, Ren_Mon_Tra,")
			loConsulta.AppendLine("        Ren_Mon_Per, Ren_Mon_Ban, Ren_Mon_Adu, Ren_Mon_Arc,")
			loConsulta.AppendLine("        Ren_Mon_Ots1, Ren_Mon_Ots2, Ren_Mon_Ots3")
			loConsulta.AppendLine("		)")
			loConsulta.AppendLine("SELECT	2 AS Tipo, R.Origen, R.Documento, R.Adicional, ")
			loConsulta.AppendLine("		Renglon, Cod_Art, Nom_Art, Cod_Alm, ")
			loConsulta.AppendLine("		Can_Art1, Precio1, Mon_Bru,")
			loConsulta.AppendLine("        Mon_Seg, Mon_Fle, Mon_Alm,")
			loConsulta.AppendLine("        Mon_Ipt, Mon_Por, Mon_Tra,")
			loConsulta.AppendLine("        Mon_Per, Mon_Ban, Mon_Adu, Mon_Arc,")
			loConsulta.AppendLine("        Mon_Ots1, Mon_Ots2, Mon_Ots3")
			loConsulta.AppendLine("FROM	Renglones_Importaciones AS R")
			loConsulta.AppendLine("	JOIN #tmpDocumentoActual")
			loConsulta.AppendLine("		ON	#tmpDocumentoActual.Documento	= R.Documento")
			loConsulta.AppendLine("		AND	#tmpDocumentoActual.Origen		= R.Origen")
			loConsulta.AppendLine("		AND	#tmpDocumentoActual.Adicional	= R.Adicional")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("/* Detalles.																				*/")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("INSERT INTO #tmpImportacion")
			loConsulta.AppendLine("	(	Tipo, Origen, Documento, Adicional, Renglon, ")
			loConsulta.AppendLine("		Det_Grupo, Det_Afe_Cos, Det_Cod_Gas, Det_Nom_Gas, Det_Mon_Gas, ")
			loConsulta.AppendLine("		Det_Tip_Ori, Det_Doc_Ori, Det_Ren_Ori, ")
			loConsulta.AppendLine("		Det_Cod_Art, Det_Nom_Art, Det_Can_Art, ")
			loConsulta.AppendLine("		Det_Precio1, Det_Cod_Con, Det_Nom_Con, ")
			loConsulta.AppendLine("		Det_Mon_Deb, Det_Mon_Hab, Det_Cod_Imp, ")
			loConsulta.AppendLine("		Det_Por_Imp, Det_Mon_Imp, Det_Mon_Bru, ")
			loConsulta.AppendLine("		Det_Mon_Net, Det_Tip_Gas, Det_Con_Gas")
			loConsulta.AppendLine("		)")
			loConsulta.AppendLine("SELECT	(CASE Grupo")
			loConsulta.AppendLine("			 WHEN 'Gastos Fijos'				THEN 3")
			loConsulta.AppendLine("			 WHEN 'Gastos Adicionales'			THEN 4")
			loConsulta.AppendLine("			 WHEN 'Compras y Pagos Asociados'	THEN 5")
			loConsulta.AppendLine("			 WHEN 'Gastos en Compras'			THEN 6")
			loConsulta.AppendLine("			 ELSE 7 --Desconocido")
			loConsulta.AppendLine("		END) AS Tipo, D.Origen, D.Documento, D.Adicional, D.Renglon, ")
			loConsulta.AppendLine("		Grupo, Afe_Cos, Cod_Gas, Nom_Gas, Mon_Gas, ")
			loConsulta.AppendLine("		Tip_Ori, Doc_Ori, Ren_Ori, ")
			loConsulta.AppendLine("		Cod_Art, Nom_Art, Can_Art, ")
			loConsulta.AppendLine("		Precio1, Cod_Con, Nom_Con, ")
			loConsulta.AppendLine("		Mon_Deb, Mon_Hab, Cod_Imp, ")
			loConsulta.AppendLine("		Por_Imp, Mon_Imp, Mon_Bru, ")
			loConsulta.AppendLine("		Mon_Net, Tipo, Concepto")
			loConsulta.AppendLine("FROM	Detalles_Importaciones AS D")
			loConsulta.AppendLine("	JOIN #tmpDocumentoActual")
			loConsulta.AppendLine("		ON	#tmpDocumentoActual.Documento	= D.Documento")
			loConsulta.AppendLine("		AND	#tmpDocumentoActual.Origen		= D.Origen")
			loConsulta.AppendLine("		AND	#tmpDocumentoActual.Adicional	= D.Adicional")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("/* Encabezado.																				*/")
			loConsulta.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loConsulta.AppendLine("UPDATE	#tmpImportacion")
			loConsulta.AppendLine("SET		Status		= Encabezado.Status, ")
			loConsulta.AppendLine("		Referencia	= Encabezado.Referencia, ")
			loConsulta.AppendLine("		Expediente	= Encabezado.Expediente, ")
			loConsulta.AppendLine("		Fec_Ini		= Encabezado.Fec_Ini, ")
			loConsulta.AppendLine("		Cod_Mon		= Encabezado.Cod_Mon, ")
			loConsulta.AppendLine("		Tasa		= Encabezado.Tasa, ")
			loConsulta.AppendLine("		Cod_Pro		= Encabezado.Cod_Pro, ")
			loConsulta.AppendLine("		Nom_Pro		= Encabezado.Nom_Pro, ")
			loConsulta.AppendLine("		Comentario	= Encabezado.Comentario, ")
			loConsulta.AppendLine("		Tip_Dis		= Encabezado.Tip_Dis, ")
			loConsulta.AppendLine("		Tip_Ara		= Encabezado.Tip_Ara, ")
			loConsulta.AppendLine("		Mon_Bru		= Encabezado.Mon_Bru, ")
			loConsulta.AppendLine("		Mon_Gas_Fij	= Encabezado.Mon_Gas_Fij, ")
			loConsulta.AppendLine("		Mon_Gas_Adi	= Encabezado.Mon_Gas_Adi, ")
			loConsulta.AppendLine("		Mon_Com_Pag	= Encabezado.Mon_Com_Pag, ")
			loConsulta.AppendLine("		Mon_Gas_Com	= Encabezado.Mon_Gas_Com, ")
			loConsulta.AppendLine("		Mon_Ara		= Encabezado.Mon_Ara, ")
			loConsulta.AppendLine("		Mon_Net		= Encabezado.Mon_Net,")
			loConsulta.AppendLine("	    Tas_Emi	    = Encabezado.Tas_Emi,")
			loConsulta.AppendLine("	    Tas_Car	    = Encabezado.Tas_Car,")
			loConsulta.AppendLine("	    Tas_Bar	    = Encabezado.Tas_Bar,")
			loConsulta.AppendLine("	    Tas_Otr1	= Encabezado.Tas_Otr1,")
			loConsulta.AppendLine("	    Tas_Otr2	= Encabezado.Tas_Otr2,")
			loConsulta.AppendLine("	    Tas_Otr3	= Encabezado.Tas_Otr3,")
			loConsulta.AppendLine("	    Tas_Otr4	= Encabezado.Tas_Otr4,")
			loConsulta.AppendLine("	    Tas_Otr5	= Encabezado.Tas_Otr5")
			loConsulta.AppendLine("		")
			loConsulta.AppendLine("FROM	(		")
			loConsulta.AppendLine("	SELECT		I.Status, I.Referencia, I.Expediente, I.Fec_Ini, I.Cod_Mon, I.Tasa, ")
			loConsulta.AppendLine("				P.Cod_Pro, P.Nom_Pro, I.Comentario, I.Tip_Dis, I.Tip_Ara, ")
			loConsulta.AppendLine("				I.Mon_Bru, I.Mon_Gas_Fij, I.Mon_Gas_Adi, I.Mon_Com_Pag, ")
			loConsulta.AppendLine("				I.Mon_Gas_Com, I.Mon_Ara, I.Mon_Net,")
			loConsulta.AppendLine("                I.Tas_Emi, I.Tas_Car, I.Tas_Bar, I.Tas_Otr1,")
			loConsulta.AppendLine("                I.Tas_Otr2, I.Tas_Otr3, I.Tas_Otr4, I.Tas_Otr5")
			loConsulta.AppendLine("	FROM		Importaciones AS I")
			loConsulta.AppendLine("		JOIN	Proveedores AS P")
			loConsulta.AppendLine("			ON	I.Cod_Pro = P.Cod_Pro")
			loConsulta.AppendLine("		JOIN	#tmpDocumentoActual")
			loConsulta.AppendLine("			ON	#tmpDocumentoActual.Documento	= I.Documento")
			loConsulta.AppendLine("			AND	#tmpDocumentoActual.Origen		= I.Origen")
			loConsulta.AppendLine("			AND	#tmpDocumentoActual.Adicional	= I.Adicional")
			loConsulta.AppendLine(") AS Encabezado;")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		*  ")
            loConsulta.AppendLine("FROM		#tmpImportacion")
			loConsulta.AppendLine("ORDER BY	Tipo, Renglon")
			loConsulta.AppendLine("")
			'loConsulta.AppendLine("DROP TABLE #tmpImportacion")
			'loConsulta.AppendLine("DROP TABLE #tmpDocumentoActual")
			loConsulta.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
			
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCostos_Importaciones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCostos_Importaciones.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 14/08/12: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 17/08/12: En Detalles_Importaciones se cambiaron las referencias a Caracter1 y		'
'				 Caracter2 por 	Tipo y Concepto, respectivamente.							'
'-------------------------------------------------------------------------------------------'
' RJG: 14/11/14: Se agregaron los gastos totales distribuidos (por categorías) y las tasas  '
'                aplicadas.         							                            '
'-------------------------------------------------------------------------------------------'
