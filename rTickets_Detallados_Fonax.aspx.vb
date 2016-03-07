'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTickets_Detallados_Fonax"
'-------------------------------------------------------------------------------------------'
Partial Class rTickets_Detallados_Fonax 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))            
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT	MAX(Auditorias.Registro) AS Fecha_Confirmacion, "  )
			loComandoSeleccionar.Appendline("		Auditorias.Documento "  )
			loComandoSeleccionar.Appendline("INTO #tmpTemporal "  )
			loComandoSeleccionar.Appendline("FROM Auditorias "  )
			loComandoSeleccionar.Appendline("WHERE 	Auditorias.Tabla = 'Cotizaciones' "  )
			loComandoSeleccionar.Appendline("		AND Auditorias.Accion = 'Confirmar' "  )
			loComandoSeleccionar.Appendline("GROUP BY Auditorias.Documento "  )
			loComandoSeleccionar.Appendline("ORDER BY Documento ASC"  )
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			loComandoSeleccionar.Appendline("SELECT MIN(Auditorias.Registro) AS Registro_Original,Auditorias.Documento"  )
			loComandoSeleccionar.Appendline("INTO #tmpTemporalVendedor"  )
			loComandoSeleccionar.Appendline("FROM Auditorias"  )
			loComandoSeleccionar.Appendline("WHERE Tabla = 'COTIZACIONES' AND Accion = 'MODIFICAR' AND Detalle LIKE '%COD_VEN%'"  )
			loComandoSeleccionar.Appendline("GROUP BY Documento "  )
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			loComandoSeleccionar.Appendline("SELECT #tmpTemporalVendedor.Documento AS Documento,Auditorias.detalle AS Detalle"  )
			loComandoSeleccionar.Appendline("INTO #tmpVendedorModificado"  )
			loComandoSeleccionar.Appendline("FROM #tmpTemporalVendedor"  )
			loComandoSeleccionar.Appendline("JOIN Auditorias ON (Auditorias.Registro = #tmpTemporalVendedor.Registro_Original AND Auditorias.Documento = #tmpTemporalVendedor.Documento)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")


			
			
			loComandoSeleccionar.Appendline("SELECT	Cotizaciones.Documento, "  )
			loComandoSeleccionar.Appendline("		ISNULL(#tmpVendedorModificado.Detalle,'') AS Detalle, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Reg_Ini AS Fecha_Creacion, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Reg_Mod AS Fecha_Modificacion, "  )
			loComandoSeleccionar.Appendline("		ISNULL(#tmpTemporal.Fecha_Confirmacion,0) AS Fecha_Confirmacion, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Ven, ")
			loComandoSeleccionar.Appendline("		Cotizaciones.Usu_Mod, ")
			loComandoSeleccionar.Appendline("		Vendedores.Nom_Ven, "  )
			loComandoSeleccionar.Appendline("		'' AS lcOperadorInicialAsignado, "  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Status,"  )
			loComandoSeleccionar.Appendline("		Cotizaciones.Cod_Cli,"  )
			loComandoSeleccionar.Appendline("		Clientes.Nom_Cli, "  )
			loComandoSeleccionar.Appendline("		Clientes.Telefonos, "  )
			loComandoSeleccionar.Appendline("		Clientes.Correo,"  )
			loComandoSeleccionar.Appendline("		Clientes.Cod_Pai,"  )
			loComandoSeleccionar.Appendline("		Clientes.Dir_Fis,"  )
			loComandoSeleccionar.Appendline("		Paises.Nom_Pai,"  )
			loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Cotizaciones.Status = 'Pendiente' THEN (DATEDIFF(Day,Cotizaciones.Reg_Ini,getDate()))")
			loComandoSeleccionar.AppendLine("			WHEN Cotizaciones.Status = 'Confirmado' THEN (DATEDIFF(Day,Cotizaciones.Reg_Ini,#tmpTemporal.Fecha_Confirmacion))")
			loComandoSeleccionar.AppendLine("			ELSE 0")
			loComandoSeleccionar.AppendLine("		END AS Dias_Activo,")
			loComandoSeleccionar.AppendLine("		Renglones_Cotizaciones.Comentario,")
			loComandoSeleccionar.AppendLine("		Renglones_Cotizaciones.Renglon,")
			loComandoSeleccionar.AppendLine("		Renglones_Cotizaciones.Cod_Art,")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine("		Renglones_Cotizaciones.Registro AS Fecha_Renglon_Agregado,")
			loComandoSeleccionar.AppendLine("		Articulos.Atributo_A")
			loComandoSeleccionar.AppendLine("FROM	Cotizaciones")
			loComandoSeleccionar.AppendLine("JOIN Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
			loComandoSeleccionar.Appendline("							And Renglones_Cotizaciones.Cod_Art between " & lcParametro5Desde )
			loComandoSeleccionar.Appendline("							And " & lcParametro5Hasta )
			
			loComandoSeleccionar.AppendLine("JOIN Clientes ON (Cotizaciones.Cod_Cli = Clientes.Cod_Cli )")
			loComandoSeleccionar.AppendLine("JOIN Vendedores ON (Cotizaciones.Cod_Ven = Vendedores.Cod_Ven)")
			loComandoSeleccionar.AppendLine("JOIN Articulos ON Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.Appendline("							And Articulos.Cod_Dep between " & lcParametro6Desde )
			loComandoSeleccionar.Appendline("							And " & lcParametro6Hasta )
			loComandoSeleccionar.Appendline("							And Articulos.Cod_Sec between " & lcParametro7Desde )
			loComandoSeleccionar.Appendline("							And " & lcParametro7Hasta )
			
			If cusAplicacion.goReportes.paParametrosIniciales(10) <> "Todos" Then
					
					loComandoSeleccionar.Appendline("					And Articulos.Atributo_A IN (" & lcParametro10Desde & ")")
			
			End If
			
			loComandoSeleccionar.AppendLine("JOIN Paises ON (Paises.Cod_Pai = Clientes.Cod_Pai)")
			loComandoSeleccionar.Appendline("LEFT JOIN #tmpVendedorModificado ON (Cotizaciones.Documento = #tmpVendedorModificado.Documento)"  )
			loComandoSeleccionar.AppendLine("LEFT JOIN #tmpTemporal ON (#tmpTemporal.Documento = Cotizaciones.Documento)")
			loComandoSeleccionar.Appendline("WHERE	Cotizaciones.Documento Between " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Fec_Ini between " & lcParametro1Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine("		And Cotizaciones.Status IN (" & lcParametro2Desde & ")")
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Cli between " & lcParametro3Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Ven between " & lcParametro4Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro4Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Tra between " & lcParametro8Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro8Hasta )
			loComandoSeleccionar.Appendline("		And Cotizaciones.Cod_Rev between " & lcParametro9Desde )
			loComandoSeleccionar.Appendline("		And " & lcParametro9Hasta )
             loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString , "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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
            
            
		 '-------------------------------------------------------------------------------------------------------
		 ' Verificando si el Código del Vendedor se ha modificado
		 '-------------------------------------------------------------------------------------------------------  
            
            
            Dim lcXml As String = "<detalle><campo></campo></detalle>"
            Dim loDetalle As New System.Xml.XmlDocument()
            Dim lcCodigoOriginal As String   = ""
            
            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Detalle")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("lcOperadorInicialAsignado") = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Nom_Ven").ToString.Trim
                    Continue For
                End If


                loDetalle.LoadXml(lcXml)

                'En cada renglón lee el contenido del campo detalle
                For Each loFilaDetalle As System.Xml.XmlNode In loDetalle.SelectNodes("detalle/campos/campo[@nombre='cod_ven']")
					
						
					If (loFilaDetalle.SelectSingleNode("antes").InnerXml)<> "" Then

							lcCodigoOriginal = (loFilaDetalle.SelectSingleNode("antes").InnerXml.Trim) 
				
					End If
								
				
                Next loFilaDetalle
               
				If	 lcCodigoOriginal <> "" Then

						Dim loComandoSeleccion  As New StringBuilder()
	            
						loComandoSeleccion.AppendLine("SELECT Nom_Ven FROM Vendedores WHERE Cod_Ven = " & goServicios.mObtenerCampoFormatoSQL(lcCodigoOriginal.Trim()))
	           
						Dim laDatosOperadorInicial As DataTable = loServicios.mObtenerTodosSinEsquema(loComandoSeleccion.ToString , "curReportes").Tables(0)
					    
					    laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("lcOperadorInicialAsignado") = laDatosOperadorInicial.Rows(0).Item("Nom_Ven").ToString.Trim
                    
					    
						
				End If 
                
            Next lnNumeroFila
            
            
            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTickets_Detallados_Fonax", laDatosReporte)
            
            
           
            
           
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrTickets_Detallados_Fonax.ReportSource =	 loObjetoReporte	

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
' MAT: 20/06/11: Código Inicial
'-------------------------------------------------------------------------------------------'
' MAT: 27/06/11: Culminación de la programación de las modificaciones según requerimientos
'-------------------------------------------------------------------------------------------'
' MAT: 21/07/11: Agregado nuevo filtro Atributo
'-------------------------------------------------------------------------------------------'