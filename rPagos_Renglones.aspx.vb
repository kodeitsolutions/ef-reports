'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPagos_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rPagos_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosFinales(8)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim llFiltrarMovimientos AS Boolean =	(lcParametro5Desde<>"''") OrElse (lcParametro6Desde<>"''") OrElse _
													(lcParametro5Hasta<>"'zzzzzzz'") OrElse (lcParametro6Hasta<>"'zzzzzzz'")

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT		Pagos.Cod_Pro						AS Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro					AS Nom_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Rif						AS Rif, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nit						AS Nit, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis					AS Dir_Fis, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Telefonos				AS Telefonos, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Fax						AS Fax, ")
			loComandoSeleccionar.AppendLine("			Pagos.Documento						AS Documento, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini						AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin						AS Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Bru						AS Mon_Bru, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Imp						AS Mon_Imp, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Net						AS Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Pagos.Comentario					AS Comentario, ")
			loComandoSeleccionar.AppendLine("			Renglones_Pagos.Renglon     		AS Ren_Doc, ")
			loComandoSeleccionar.AppendLine("			Renglones_Pagos.Tip_Doc     		AS Tip_Doc, ")
			loComandoSeleccionar.AppendLine("			Renglones_Pagos.Cod_Tip     		AS Cod_Tip, ")
			loComandoSeleccionar.AppendLine("			Renglones_Pagos.Doc_Ori     		AS Doc_Ori, ")
			loComandoSeleccionar.AppendLine("			(CASE ")
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
			loComandoSeleccionar.AppendLine("				THEN Renglones_Pagos.Mon_Net ")
			loComandoSeleccionar.AppendLine("				ELSE -Renglones_Pagos.Mon_Net ")
			loComandoSeleccionar.AppendLine("			END)								AS  Mon_NetD, ")
			loComandoSeleccionar.AppendLine("			(CASE ")
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
			loComandoSeleccionar.AppendLine("				THEN Renglones_Pagos.Mon_Abo ")
			loComandoSeleccionar.AppendLine("				ELSE -Renglones_Pagos.Mon_Abo")
			loComandoSeleccionar.AppendLine("			END)								AS  Mon_Abo, ")
			loComandoSeleccionar.AppendLine("			0.00                        		AS  Ren_Tip, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Tip_Ope, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Doc_Des, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Num_Doc, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Cod_Caj, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Cod_Ban, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			SPACE(10)                   		AS  Cod_Tar, ")
			loComandoSeleccionar.AppendLine("			0.00                        		AS  Mon_NetTP, ")
			loComandoSeleccionar.AppendLine("			'Documentos'                		AS  Tipo, ")
			loComandoSeleccionar.AppendLine("			SPACE(25)                   		AS  Nom_Caj, ")
			loComandoSeleccionar.AppendLine("			SPACE(25)                   		AS  Nom_Ban, ")
			loComandoSeleccionar.AppendLine("			SPACE(25)                   		AS  Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			SPACE(25)                   		AS  Nom_Tar ")
			loComandoSeleccionar.AppendLine("INTO		#tmpDocumentos ")
			loComandoSeleccionar.AppendLine("FROM		Pagos  ")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Pagos ")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Pagos.Documento = Pagos.Documento")
			loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Pagar")
			loComandoSeleccionar.AppendLine("		ON	Cuentas_Pagar.Cod_Tip = Renglones_Pagos.Cod_Tip ")
			loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori ")
			loComandoSeleccionar.AppendLine("	JOIN	Proveedores ")
			loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Pagos.Cod_Pro")
			loComandoSeleccionar.AppendLine("WHERE		Pagos.Documento	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("       AND Pagos.Fec_Ini	BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Pro	BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("       AND Pagos.Status	IN (" & lcParametro3Desde &")")
			loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Suc	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			    
			If lcParametro8Desde = "Igual" Then
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev	BETWEEN " & lcParametro7Desde & "AND " & lcParametro7Hasta)
			Else
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev	NOT BETWEEN " & lcParametro7Desde & "AND " & lcParametro7Hasta)
			End If
            
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		Pagos.Cod_Pro							AS Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro						AS Nom_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Rif							AS Rif, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nit							AS Nit, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis						AS Dir_Fis, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Telefonos					AS Telefonos, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Fax							AS Fax, ")
			loComandoSeleccionar.AppendLine("			Pagos.Documento							AS Documento, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini							AS Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin							AS Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Bru							AS Mon_Bru, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Imp							AS Mon_Imp, ")
			loComandoSeleccionar.AppendLine("			Pagos.Mon_Net							AS Mon_Net, ")
			loComandoSeleccionar.AppendLine("			Pagos.Comentario						AS Comentario, ")
			loComandoSeleccionar.AppendLine("			0                           			AS Ren_Doc, ")
			loComandoSeleccionar.AppendLine("			''                          			AS Tip_Doc, ")
			loComandoSeleccionar.AppendLine("			''                          			AS Cod_Tip, ")
			loComandoSeleccionar.AppendLine("			''                          			AS Doc_Ori, ")
			loComandoSeleccionar.AppendLine("			0.00                        			AS Mon_NetD, ")
			loComandoSeleccionar.AppendLine("			0.00                        			AS Mon_Abo, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Renglon      			AS Ren_Tip, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Tip_Ope      			AS Tip_Ope, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Doc_Des      			AS Doc_Des, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Num_Doc      			AS Num_Doc, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Caj      			AS Cod_Caj, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Ban      			AS Cod_Ban, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Cue      			AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Cod_Tar      			AS Cod_Tar, ")
			loComandoSeleccionar.AppendLine("			Detalles_Pagos.Mon_Net      			AS Mon_NetTP, ")
			loComandoSeleccionar.AppendLine("			'TiposPagos'                			AS Tipo, ")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Cajas.Nom_Caj,1,25),'')			AS Nom_Caj, ")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Bancos.Nom_Ban,1,25),'')			AS Nom_Ban, ")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25),'')AS Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			ISNULL(SUBSTRING(Tarjetas.Nom_Tar,1,25),'')			AS Nom_Tar ")
			loComandoSeleccionar.AppendLine("INTO		#tmpTiposPagos ")
			loComandoSeleccionar.AppendLine("FROM		Pagos ")
			loComandoSeleccionar.AppendLine("	JOIN	Detalles_Pagos")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Documento = Pagos.Documento")
			loComandoSeleccionar.AppendLine("	JOIN	Proveedores ")
			loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Pagos.Cod_Pro")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Cajas ")
			loComandoSeleccionar.AppendLine("		ON  Detalles_Pagos.Cod_Caj = Cajas.Cod_Caj ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Bancos ")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Cod_Ban = Bancos.Cod_Ban ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Bancarias ")
			loComandoSeleccionar.AppendLine("		ON	Detalles_Pagos.Cod_Cue = Cuentas_Bancarias.Cod_Cue ")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Tarjetas ")
			loComandoSeleccionar.AppendLine("		ON  Detalles_Pagos.Cod_Tar = Tarjetas.Cod_Tar ")
			loComandoSeleccionar.AppendLine("WHERE		Pagos.Documento			BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Fec_Ini			BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Pro			BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND Pagos.Status			IN (" & lcParametro3Desde &")")
			loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Suc			BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("		AND Detalles_Pagos.Cod_Caj	BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("		AND Detalles_Pagos.Cod_Cue	BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
  
			If lcParametro8Desde = "Igual" Then
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			Else
			    loComandoSeleccionar.AppendLine("		AND Pagos.Cod_rev 			NOT BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			End If
          
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("	SELECT * FROM #tmpDocumentos")
			If llFiltrarMovimientos Then 
				loComandoSeleccionar.AppendLine("	WHERE #tmpDocumentos.Documento IN (SELECT Documento FROM #tmpTiposPagos)")
			End If
			loComandoSeleccionar.AppendLine("UNION ALL  ")
			loComandoSeleccionar.AppendLine("	SELECT * FROM #tmpTiposPagos")
			loComandoSeleccionar.AppendLine("ORDER BY 8, " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			    
            

            Dim loServicios As New cusDatos.goDatos
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_Renglones", laDatosReporte)
			
						
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Renglones.ReportSource = loObjetoReporte

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
' JJD: 14/03/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' GCR: 30/03/09: Ajustes al diseño															'
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"														'
'-------------------------------------------------------------------------------------------'
' CMS:  10/08/09: Metodo de ordenamiento, verificacionde registros							'
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/10: Filtro Caja, Cuenta														'
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó total de renglones y documentos.								'
'-------------------------------------------------------------------------------------------'
' RJG:  27/04/12: Se ajustó el código y se simplificó el SELECT (por medio de JOINS).		'
'-------------------------------------------------------------------------------------------'
' RJG:  13/10/12: Se corrigió el filtro de Caja/Cuenta: no filtraba los encabezados de los	'
'				  pagos asociados a ellos, solo filtraba los detalles.						'
'-------------------------------------------------------------------------------------------'
