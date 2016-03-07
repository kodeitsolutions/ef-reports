'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Compras
    Inherits vis2formularios.frmReporte
    
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


	Try	
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)
			Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosFinales(6)
			Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	ROW_NUMBER() OVER (Partition by Cuentas_Pagar.Cod_Tip ORDER BY  " & lcOrdenamiento &" ) AS 'Renglon'," ) 
			'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tip, " )
			loComandoSeleccionar.AppendLine("			 CASE" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'FACT' THEN 'Factura'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'GIRO' THEN 'Giro'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'ISRL' THEN 'ISRL'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'N/CR' THEN 'Nota de Credito'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'N/DB' THEN 'Nota de Debito'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'RETIVA' THEN 'Retensión de I.V.A.'" )
			loComandoSeleccionar.AppendLine("			 END AS cod_tip," )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc, " )
			'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Bru, " )
			'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Net, " )
			loComandoSeleccionar.AppendLine("			CASE  " )
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Bru * -1 " )
			loComandoSeleccionar.AppendLine("				ELSE  " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Bru " )
			loComandoSeleccionar.AppendLine("			END AS Mon_Bru,  " )
			loComandoSeleccionar.AppendLine("			 " )
			loComandoSeleccionar.AppendLine("			CASE  " )
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Net * -1 " )
			loComandoSeleccionar.AppendLine("				ELSE  " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Net " )
			loComandoSeleccionar.AppendLine("			END AS Mon_Net,  " )

			'loComandoSeleccionar.AppendLine("Cuentas_Pagar.Mon_Imp1, " )
			'loComandoSeleccionar.AppendLine("Cuentas_Pagar.Mon_bas1 as Mon_bas, " )
			'loComandoSeleccionar.AppendLine("Cuentas_Pagar.por_imp1, " )
			'loComandoSeleccionar.AppendLine("Cuentas_Pagar.Mon_exe, " )

			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.dis_imp, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe1, " )

			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe2, " )

			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe3, " )
				
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_exe, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_bas, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_imp, " )

			'loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, " )
			
			loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) END) AS  Rif, ")
            			
			loComandoSeleccionar.AppendLine("			Proveedores.tip_Pro, " )
			'loComandoSeleccionar.AppendLine("			Proveedores.Rif, " )
			loComandoSeleccionar.AppendLine("			(Case When Cuentas_Pagar.Status = 'Anulado' Then '03-ANU' Else '01-REG' End) as Transaccion " )
			loComandoSeleccionar.AppendLine("INTO		#tempCxP " )
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar, Proveedores " )
			loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro " )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
			
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT between " & lcParametro4Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			
			
			loComandoSeleccionar.AppendLine(" 			And (Cuentas_Pagar.cod_tip = 'FACT' OR Cuentas_Pagar.cod_tip = 'GIRO' OR Cuentas_Pagar.cod_tip = 'ISRL' OR ")
			loComandoSeleccionar.AppendLine(" 			Cuentas_Pagar.cod_tip = 'N/CR' OR Cuentas_Pagar.cod_tip = 'N/DB' OR Cuentas_Pagar.cod_tip = 'RETIVA')")
			loComandoSeleccionar.AppendLine("ORDER BY   Cuentas_Pagar.Cod_Tip,  " & lcOrdenamiento)
			
			
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''Obtencion de las Ordenes de pago '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			
			loComandoSeleccionar.AppendLine(" SELECT	ROW_NUMBER() OVER (ORDER BY  " & lcOrdenamiento.Replace("Cuentas_Pagar.", "Ordenes_Pagos.") &" ) AS 'Renglon'," ) 
			loComandoSeleccionar.AppendLine("			'Orden de Pago' AS cod_tip, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Documento, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Control, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Factura, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("			'debito' As Tip_Doc, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Bru, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net, " )

			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.dis_imp, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe1, " )

			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe2, " )

			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe3, " )
				
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_exe, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_bas, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_imp, " )
			
			loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif, ")
            			
			loComandoSeleccionar.AppendLine("			Proveedores.tip_Pro, " )
			loComandoSeleccionar.AppendLine("			(Case When Ordenes_Pagos.Status = 'Confirmado' Then '01-REG' Else '03-ANU' End) as Transaccion " )
			loComandoSeleccionar.AppendLine("INTO		#tempOrdenes_Pago " )
			loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos, Proveedores " )
			loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro " )
			
			If lcParametro7Desde.ToUpper = "NO" Then

				loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Mon_Imp <> 0" )
				
			End If
			
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
			
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Ordenes_Pagos.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Ordenes_Pagos.Cod_Rev NOT between " & lcParametro4Desde)
            End If

            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			
			
			If lcParametro6Desde.ToUpper = "SI" then 

				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempCxP	")
				loComandoSeleccionar.AppendLine(" UNION ALL	")
				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempOrdenes_Pago	")
				loComandoSeleccionar.AppendLine("ORDER BY   Cod_Tip,  " & lcOrdenamiento.Replace("Cuentas_Pagar.", " "))
				
			Else
			
				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempCxP ")
			
			End If

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

	        
	        Dim loServicios As New cusDatos.goDatos
			
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			
			
			Dim loDistribucion	As System.Xml.XmlDocument
			Dim laImpuestos		As system.Xml.XmlNodeList
			
			For Each loFila As DataRow In laDatosReporte.Tables(0).Rows
				
				If Not String.IsNullOrEmpty(Trim(loFila.Item("dis_imp"))) Then 
						
						loDistribucion = New System.Xml.XmlDocument()
					Try	

						loDistribucion.LoadXml(Trim(loFila.Item("dis_imp")))
						
					Catch ex As Exception
						
						Continue For
						
					End Try
					'tuXML.firstChild.childNodes.length 
					'trace( tuXML.firstChild.childNodes.length ) 
					'trace( tuXML.firstChild.childNodes[0].childNodes.length ) 
					'trace( tuXML.firstChild.childNodes[0].childNodes[0].length ) 
						
						laImpuestos = loDistribucion.SelectNodes("impuestos/impuesto")
						
						If laImpuestos.Count >= 1 Then 
							If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
								loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText) 
								loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
								loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
								loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
							Else
								loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
								loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
								loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
								loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
							End If
						End If
						
						If laImpuestos.Count >= 2 Then 
							If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
								loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText) 
								loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
								loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
								loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
							Else
								loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
								loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
								loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
								loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
							End If
						End If
						
						If laImpuestos.Count >= 3 Then 
							If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
								loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText) 
								loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
								loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
								loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
							Else
								loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
								loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
								loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
								loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
							End If
						End If

							loFila.Item("subt_imp") = loFila.Item("mon_imp3") + loFila.Item("mon_imp2") + loFila.Item("mon_imp1")
							loFila.Item("subt_exe") = loFila.Item("mon_exe1") + loFila.Item("mon_exe2") + loFila.Item("mon_exe3")
							loFila.Item("subt_bas") = loFila.Item("mon_bas1") + loFila.Item("mon_bas2") + loFila.Item("mon_bas3")

				End If
				
			Next lofila
 
 
			laDatosReporte.Tables(0).AcceptChanges()

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

 
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
	   
			Me.crvrLibro_Compras.ReportSource = loObjetoReporte
			
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
		
			 loObjetoReporte.Close ()

		Catch loExcepcion As Exception

		End Try
	
	End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 02/08/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' YJP: 13/05/09: Ajustes en el Reporte														'
'-------------------------------------------------------------------------------------------'
' AAP: 01/07/09: Filtro "Sucursal:"															'
'-------------------------------------------------------------------------------------------'
' CMS: 03/04/10: Se ajusto para tomar los siguientes documentos GIRO, ISRL, N/CR, N/DB y	'
'				 RETIVA.																	'
'-------------------------------------------------------------------------------------------'
' CMS: 03/04/10: Se ajusto para considerar negativos los documentos tipo credito			'
'					Validacion de registro cero												'
'-------------------------------------------------------------------------------------------'
' CMS: 18/05/10: Se ajusto para tomar el proveedor generico									'
'-------------------------------------------------------------------------------------------'
' CMS: 22/05/10: Filtro Revision y Tipo de revision.										'
'-------------------------------------------------------------------------------------------'
' CMS: 17/05/10: Se a gregaron los filtros ¿Ordenes de Pago? y ¿Ordenes de Pago Exentas?.	'
'				 lo que conllevó a ampliar la consulta para obtener y unir las ordenes de	'
'				 pagos a la consulta ya existente.											'
'-------------------------------------------------------------------------------------------'
' RJG: 21/02/13: Se corrigió el cálculo de los totales, para no incluir las CXC anuladas.	'
'-------------------------------------------------------------------------------------------'
' RJG: 30/08/13: Se corrigió el cálculo de los totales para no incluir las Ordenes de Pago  '
'                pendientes.                                                            	'
'-------------------------------------------------------------------------------------------'
