'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Compras_Lirio"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Compras_Lirio
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
			Dim loConsulta As New StringBuilder()

			loConsulta.AppendLine("") 
			loConsulta.AppendLine("SELECT      ROW_NUMBER() OVER (Partition by Cuentas_Pagar.Cod_Tip ORDER BY " & lcOrdenamiento & ") AS Renglon,") 
			loConsulta.AppendLine("            CASE") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'FACT' THEN 'Factura'") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'GIRO' THEN 'Giro'") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'ISRL' THEN 'ISRL'") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'N/CR' THEN 'Nota de Credito'") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'N/DB' THEN 'Nota de Debito'") 
			loConsulta.AppendLine("                WHEN Cuentas_Pagar.cod_tip = 'RETIVA' THEN 'Retensión de I.V.A.'") 
			loConsulta.AppendLine("                ELSE Cuentas_Pagar.cod_tip") 
			loConsulta.AppendLine("            END AS cod_tip,") 
			loConsulta.AppendLine("            Cuentas_Pagar.Documento, ") 
			loConsulta.AppendLine("            Cuentas_Pagar.Control, ") 
			loConsulta.AppendLine("            Cuentas_Pagar.Factura, ") 
			loConsulta.AppendLine("            Cuentas_Pagar.Cod_Pro, ") 
			loConsulta.AppendLine("            Cuentas_Pagar.Fec_Ini, ") 
			loConsulta.AppendLine("            Cuentas_Pagar.Tip_Doc, ") 
			loConsulta.AppendLine("            CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ") 
			loConsulta.AppendLine("                THEN Cuentas_Pagar.Mon_Bru * -1 ") 
			loConsulta.AppendLine("                ELSE Cuentas_Pagar.Mon_Bru ") 
			loConsulta.AppendLine("            END AS Mon_Bru,  ") 
			loConsulta.AppendLine("            CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ") 
			loConsulta.AppendLine("                THEN Cuentas_Pagar.Mon_Net * -1 ") 
			loConsulta.AppendLine("                ELSE Cuentas_Pagar.Mon_Net ") 
			loConsulta.AppendLine("            END AS Mon_Net,  ") 
			loConsulta.AppendLine("            Cuentas_Pagar.dis_imp, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_imp1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_bas1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS por_imp1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_exe1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_imp2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_bas2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS por_imp2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_exe2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_imp3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_bas3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS por_imp3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS mon_exe3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS subt_exe, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS subt_bas, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) AS subt_imp, ") 
			loConsulta.AppendLine("            (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) AS  Nom_Pro, ") 
			loConsulta.AppendLine("            (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) AS  Rif, ") 
			loConsulta.AppendLine("            Proveedores.tip_Pro, ")
			loConsulta.AppendLine("            (CASE WHEN Cuentas_Pagar.Status = 'Anulado' THEN '03-ANU' ELSE '01-REG' END) AS Transaccion ") 
			loConsulta.AppendLine("INTO        #tempDocumentos ") 
			loConsulta.AppendLine("FROM        Cuentas_Pagar") 
			loConsulta.AppendLine("    JOIN    Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro ") 
			loConsulta.AppendLine("WHERE       Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde )
			loConsulta.AppendLine("            AND " & lcParametro0Hasta )
			loConsulta.AppendLine("            AND Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde )
			loConsulta.AppendLine("            AND " & lcParametro1Hasta) 
			loConsulta.AppendLine("            AND Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde )
			loConsulta.AppendLine("            AND " & lcParametro2Hasta)
			loConsulta.AppendLine("            AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde )
			loConsulta.AppendLine("            AND " & lcParametro3Hasta) 
		If lcParametro5Desde = "Igual" Then
		    loConsulta.AppendLine("            AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
        Else
		    loConsulta.AppendLine("            AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
        End If
			loConsulta.AppendLine("            AND " & lcParametro4Hasta) 
		
	    '*********************************************************************
	    ' Obtencion de las Ordenes de pago (si se indica por parámetros)
	    '*********************************************************************
		If lcParametro6Desde.ToUpper = "SI" Then 

			loConsulta.AppendLine("") 
			loConsulta.AppendLine("UNION ALL") 
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      ROW_NUMBER() OVER (ORDER BY  " & lcOrdenamiento.Replace("Cuentas_Pagar.", "Ordenes_Pagos.") &" ) AS Renglon," ) 
			loConsulta.AppendLine("            'Orden de Pago' AS cod_tip, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Documento, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Control, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Factura, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Cod_Pro, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Fec_Ini, ") 
			loConsulta.AppendLine("            'debito' As Tip_Doc, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Mon_Bru, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.Mon_Net, ") 
			loConsulta.AppendLine("            Ordenes_Pagos.dis_imp, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_imp1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_bas1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As por_imp1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_exe1, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_imp2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_bas2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As por_imp2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_exe2, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_imp3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_bas3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As por_imp3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As mon_exe3, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As subt_exe, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As subt_bas, ") 
			loConsulta.AppendLine("            CAST(0.0 AS DECIMAL) As subt_imp, ") 
			loConsulta.AppendLine("            (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) AS  Nom_Pro, ") 
			loConsulta.AppendLine("            (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) AS  Rif, ") 
			loConsulta.AppendLine("            Proveedores.tip_Pro, ") 
			loConsulta.AppendLine("            (Case When Ordenes_Pagos.Status = 'Confirmado' Then '01-REG' Else '03-ANU' End) AS Transaccion ") 
			loConsulta.AppendLine("FROM        Ordenes_Pagos") 
			loConsulta.AppendLine("    JOIN    Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro") 
			loConsulta.AppendLine("WHERE       Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde )
			loConsulta.AppendLine("            AND " & lcParametro0Hasta )
			loConsulta.AppendLine("            AND Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde )
			loConsulta.AppendLine("            AND " & lcParametro1Hasta) 
			loConsulta.AppendLine("            AND Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde )
			loConsulta.AppendLine("            AND " & lcParametro2Hasta)
			loConsulta.AppendLine("            AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde )
			loConsulta.AppendLine("            AND " & lcParametro3Hasta) 
		If lcParametro5Desde = "Igual" Then
		    loConsulta.AppendLine("            AND Ordenes_Pagos.Cod_Rev BETWEEN " & lcParametro4Desde)
        Else
		    loConsulta.AppendLine("            AND Ordenes_Pagos.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
        End If
			loConsulta.AppendLine("            AND " & lcParametro4Hasta) 
			
		If lcParametro7Desde.ToUpper = "NO" Then
			loConsulta.AppendLine("            AND Ordenes_Pagos.Mon_Imp <> 0" )
		End If
			loConsulta.AppendLine("") 

        End If

            '
		    loConsulta.AppendLine("SELECT  * ") 
		    loConsulta.AppendLine("FROM    #tempDocumentos") 
		    loConsulta.AppendLine("ORDER BY   Cod_Tip, " & lcOrdenamiento.Replace("Cuentas_Pagar.", " "))
		    loConsulta.AppendLine("") 
		    loConsulta.AppendLine("") 

		    Me.mEscribirConsulta(loConsulta.ToString)

	        
	        Dim loServicios As New cusDatos.goDatos
			
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
			
			
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

 
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras_Lirio", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
	   
			Me.crvrLibro_Compras_Lirio.ReportSource = loObjetoReporte
			
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
' RJG: 27/02/15: Codigo inicial, a partir de rLibro_Compras									'
'-------------------------------------------------------------------------------------------'
