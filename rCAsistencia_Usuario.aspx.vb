Imports System.Data
Imports cusAplicacion

Partial Class rAuditorias_Usuarios_Ipos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcParametro2Desde AS String = cusAplicacion.goReportes.paParametrosIniciales(2)
			Dim lcParametro3Desde AS String = cusAplicacion.goReportes.paParametrosIniciales(3)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()
            
            
            Dim ldFecha				As DateTime	 = cusAplicacion.goReportes.paParametrosIniciales(0)
            Dim ldFechaFinal		As DateTime	= cusAplicacion.goReportes.paParametrosFinales(0)
			
			'************ INSTRUCCION PARA CREAR LA TABLA DE CADA AÑO, MES *************************
			loComandoSeleccionar.AppendLine("CREATE TABLE #tmpDias(Año INT, Mes INT, Dia INT)")
			
			If	 lcParametro2Desde = "No" AND lcParametro3Desde = "No" Then
				If ldFecha.DayOfWeek <> DayOfWeek.Saturday AND ldFecha.DayOfWeek <> DayOfWeek.Sunday Then
				
					loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
	
				End If
			Else
				 If	 lcParametro2Desde = "Si" AND lcParametro3Desde = "No" Then
					 If  ldFecha.DayOfWeek <> DayOfWeek.Sunday Then
				
						loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
				     End If
				Else
						If	 lcParametro2Desde = "No" AND lcParametro3Desde = "Si" Then
							 If  ldFecha.DayOfWeek <> DayOfWeek.Saturday Then
						
								loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
							 End If
						Else
						
							loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
		 
					    End If

				 End If
			  
			End If
			
			While   ldFecha <   ldFechaFinal
			
					ldFecha   = ldFecha.AddDays(1)
					
					If	 lcParametro2Desde = "No" AND lcParametro3Desde = "No" Then
							If ldFecha.DayOfWeek <> DayOfWeek.Saturday AND ldFecha.DayOfWeek <> DayOfWeek.Sunday Then
									loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
							End If
					Else
						If	 lcParametro2Desde = "Si" AND lcParametro3Desde = "No" Then
								 If  ldFecha.DayOfWeek <> DayOfWeek.Sunday Then
									loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
								 End If
						Else
							If	 lcParametro2Desde = "No" AND lcParametro3Desde = "Si" Then
									 If  ldFecha.DayOfWeek <> DayOfWeek.Saturday Then
										loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
									 End If
							Else
								loComandoSeleccionar.AppendLine("INSERT INTO #tmpDias VALUES(" & ldFecha.Year &  " , " & ldFecha.Month & " , " & ldFecha.Day & ")")		
			 
							End If

						End If
				  
					End If
			End While


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--************************* ENCUENTRA LAS DOS PRIMERAS FECHAS DE ENTRADA DEL USUARIO****************  ")
			loComandoSeleccionar.AppendLine("SELECT")	
			loComandoSeleccionar.AppendLine("		Auditorias.Cod_Usu,")
			loComandoSeleccionar.AppendLine("		Auditorias.Registro, ")
			loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER(PARTITION BY Auditorias.Cod_Usu, ")
			loComandoSeleccionar.AppendLine("			DATEPART(YEAR,Registro),DATEPART(MONTH,Registro),")
			loComandoSeleccionar.AppendLine("			DATEPART(DAY,Registro) ORDER BY Auditorias.Cod_Usu ASC, Auditorias.Registro) AS Item")
			loComandoSeleccionar.AppendLine("INTO #tmpTemporal_Entrada  ")
			loComandoSeleccionar.AppendLine("FROM Auditorias	")
			loComandoSeleccionar.AppendLine("WHERE Tabla = 'Usuarios' AND (ACCION = 'Acceso') AND Tipo='Seguimiento'")
			loComandoSeleccionar.AppendLine("			AND Registro Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Usu Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--************************* ENCUENTRA LAS DOS PRIMERAS FECHAS DE SALIDA DEL USUARIO****************  ")
			loComandoSeleccionar.AppendLine("SELECT")	
			loComandoSeleccionar.AppendLine("		Auditorias.Cod_Usu,")
			loComandoSeleccionar.AppendLine("		Auditorias.Registro, ")
			loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER(PARTITION BY Auditorias.Cod_Usu, ")
			loComandoSeleccionar.AppendLine("			DATEPART(YEAR,Registro),DATEPART(MONTH,Registro),")
			loComandoSeleccionar.AppendLine("			DATEPART(DAY,Registro) ORDER BY Auditorias.Cod_Usu ASC, Auditorias.Registro) AS Item")
			loComandoSeleccionar.AppendLine("INTO #tmpTemporal_Salida  ")
			loComandoSeleccionar.AppendLine("FROM Auditorias	")
			loComandoSeleccionar.AppendLine("WHERE Tabla = 'Usuarios' AND (ACCION = 'Salida') AND Tipo='Seguimiento'")
			loComandoSeleccionar.AppendLine("			AND Registro Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Usu Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal_Entrada.Cod_Usu,")
			loComandoSeleccionar.AppendLine("			DATEPART(YEAR,#tmpTemporal_Entrada.Registro)	AS Año,")
			loComandoSeleccionar.AppendLine("			DATEPART(MONTH,#tmpTemporal_Entrada.Registro)	AS Mes,  ")
			loComandoSeleccionar.AppendLine("			DATEPART(DAY,#tmpTemporal_Entrada.Registro)		AS Dia,")
			loComandoSeleccionar.AppendLine("			CASE   ")
			loComandoSeleccionar.AppendLine("					WHEN #tmpTemporal_Entrada.Item = '1' THEN #tmpTemporal_Entrada.Registro ELSE '1900-01-01 00:00:00.000'")
			loComandoSeleccionar.AppendLine("			END AS Fecha_Entrada,   ")
			loComandoSeleccionar.AppendLine("			'1900-01-01 00:00:00.000'						AS Fecha_Salida_Almuerzo, ")
			loComandoSeleccionar.AppendLine("			CASE  ")
			loComandoSeleccionar.AppendLine("					WHEN #tmpTemporal_Entrada.Item = '2' THEN #tmpTemporal_Entrada.Registro ELSE '1900-01-01 00:00:00.000'   ")
			loComandoSeleccionar.AppendLine("			END AS Fecha_Entrada_Almuerzo,")
			loComandoSeleccionar.AppendLine("			'1900-01-01 00:00:00.000'						AS Fecha_Salida	")
			loComandoSeleccionar.AppendLine("INTO	#tmpTemporal_Ent  ")
			loComandoSeleccionar.AppendLine("FROM    #tmpTemporal_Entrada ")
			loComandoSeleccionar.AppendLine("WHERE	#tmpTemporal_Entrada.Item <= 2   ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal_Salida.Cod_Usu,")
			loComandoSeleccionar.AppendLine("			DATEPART(YEAR,#tmpTemporal_Salida.Registro) AS Año,	")
			loComandoSeleccionar.AppendLine("			DATEPART(MONTH,#tmpTemporal_Salida.Registro)AS Mes,")
			loComandoSeleccionar.AppendLine("			DATEPART(DAY,#tmpTemporal_Salida.Registro)	AS Dia,")
			loComandoSeleccionar.AppendLine("			'1900-01-01 00:00:00.000' AS Fecha_Entrada,")
			loComandoSeleccionar.AppendLine("			CASE ")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal_Salida.Item = '1' THEN #tmpTemporal_Salida.Registro ELSE '1900-01-01 00:00:00.000' ")
			loComandoSeleccionar.AppendLine("			END AS Fecha_Salida_Almuerzo,")
			loComandoSeleccionar.AppendLine("			'1900-01-01 00:00:00.000' AS Fecha_Entrada_Almuerzo, ")
			loComandoSeleccionar.AppendLine("			CASE  ")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal_Salida.Item = '2' THEN #tmpTemporal_Salida.Registro ELSE '1900-01-01 00:00:00.000' ")
			loComandoSeleccionar.AppendLine("			END AS Fecha_Salida")
			loComandoSeleccionar.AppendLine("INTO	#tmpTemporal_Sal")
			loComandoSeleccionar.AppendLine("FROM    #tmpTemporal_Salida ")
			loComandoSeleccionar.AppendLine("WHERE	#tmpTemporal_Salida.Item <= 2   ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal_Entrada")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal_Salida")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT #tmpTemporal_Ent.*")
			loComandoSeleccionar.AppendLine("INTO #tmpFinal	 ")
			loComandoSeleccionar.AppendLine("FROM #tmpTemporal_Ent")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("UNION ALL ") 
			loComandoSeleccionar.AppendLine("  ")
			loComandoSeleccionar.AppendLine("SELECT #tmpTemporal_Sal.*  ")
			loComandoSeleccionar.AppendLine("FROM #tmpTemporal_Sal   ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal_Ent ")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTemporal_Sal")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			loComandoSeleccionar.AppendLine("/*******************************************************************************************************************************/   ")
			loComandoSeleccionar.AppendLine("	")
			loComandoSeleccionar.AppendLine("SELECT	#tmpFinal.Cod_Usu,")
			loComandoSeleccionar.AppendLine("		Factory_Global.dbo.Usuarios.Nom_Usu, ")
			loComandoSeleccionar.AppendLine("		#tmpFinal.Año,")
			loComandoSeleccionar.AppendLine("		#tmpFinal.Mes,")
			loComandoSeleccionar.AppendLine("		#tmpFinal.Dia,") 
			loComandoSeleccionar.AppendLine("		MAX(#tmpFinal.Fecha_Entrada) AS Fecha_Entrada,")
			loComandoSeleccionar.AppendLine("		MAX(#tmpFinal.Fecha_Salida_Almuerzo) AS Fecha_Salida_Almuerzo,")
			loComandoSeleccionar.AppendLine("		MAX(#tmpFinal.Fecha_Entrada_Almuerzo) AS Fecha_Entrada_Almuerzo,")
			loComandoSeleccionar.AppendLine("		MAX(#tmpFinal.Fecha_Salida) AS Fecha_Salida	")
			loComandoSeleccionar.AppendLine("INTO #tmpTablaFinal")
			loComandoSeleccionar.AppendLine("FROM	#tmpFinal,Factory_Global.dbo.Usuarios")
			loComandoSeleccionar.AppendLine("WHERE	(Factory_Global.dbo.Usuarios.Cod_Usu COLLATE Modern_Spanish_CI_AS = #tmpFinal.Cod_Usu COLLATE  Modern_Spanish_CI_AS)")
			loComandoSeleccionar.AppendLine("GROUP BY #tmpFinal.Cod_Usu,Factory_Global.dbo.Usuarios.Nom_Usu,#tmpFinal.Año,#tmpFinal.Mes,#tmpFinal.Dia  ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("/*******************************************************************************************************************************/")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpFinal ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		#tmpTablaFinal.Cod_Usu,")
			loComandoSeleccionar.AppendLine("			#tmpTablaFinal.Nom_Usu")
			loComandoSeleccionar.AppendLine("INTO #tmpTablaUsuarios	")
			loComandoSeleccionar.AppendLine("FROM #tmpTablaFinal ")
			loComandoSeleccionar.AppendLine("GROUP BY #tmpTablaFinal.Cod_Usu,	#tmpTablaFinal.Nom_Usu")
			loComandoSeleccionar.AppendLine("   ")
			loComandoSeleccionar.AppendLine("   ")
			loComandoSeleccionar.AppendLine("   ")
			loComandoSeleccionar.AppendLine("   ")
			loComandoSeleccionar.AppendLine("SELECT  #tmpTablaUsuarios.*,  ")
			loComandoSeleccionar.AppendLine("		 #tmpDias.*   ")
			loComandoSeleccionar.AppendLine("INTO #tmpTablaTemporal")
			loComandoSeleccionar.AppendLine("FROM #tmpDias CROSS JOIN #tmpTablaUsuarios")
			loComandoSeleccionar.AppendLine("   ")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpDias ")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTablaUsuarios")
			loComandoSeleccionar.AppendLine("/*******************************************************************************************************************************/")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT ")
			loComandoSeleccionar.AppendLine("		#tmpTablaTemporal.Cod_Usu,")		
			loComandoSeleccionar.AppendLine("		#tmpTablaTemporal.Nom_Usu,")		
			loComandoSeleccionar.AppendLine("		#tmpTablaTemporal.Año,")							
			loComandoSeleccionar.AppendLine("		#tmpTablaTemporal.Mes,")							
			loComandoSeleccionar.AppendLine("		#tmpTablaTemporal.Dia,")							
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpTablaFinal.Fecha_Entrada,'1900-01-01 00:00:00.000')			AS Fecha_Entrada,	")		
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpTablaFinal.Fecha_Salida_Almuerzo,'1900-01-01 00:00:00.000')	AS Fecha_Salida_Almuerzo,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpTablaFinal.Fecha_Entrada_Almuerzo,'1900-01-01 00:00:00.000')	AS Fecha_Entrada_Almuerzo,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpTablaFinal.Fecha_Salida,'1900-01-01 00:00:00.000')			AS Fecha_Salida	")
			loComandoSeleccionar.AppendLine("INTO #tmpTablaResultados  ")
			loComandoSeleccionar.AppendLine("FROM #tmpTablaTemporal ")
			loComandoSeleccionar.AppendLine("FULL JOIN #tmpTablaFinal ON  #tmpTablaFinal.Cod_Usu = #tmpTablaTemporal.Cod_Usu ")
			loComandoSeleccionar.AppendLine("						AND #tmpTablaFinal.Año = #tmpTablaTemporal.Año")
			loComandoSeleccionar.AppendLine("						AND #tmpTablaFinal.Mes = #tmpTablaTemporal.Mes ")
			loComandoSeleccionar.AppendLine("						AND #tmpTablaFinal.Dia = #tmpTablaTemporal.Dia ")
			loComandoSeleccionar.AppendLine("  ")
			loComandoSeleccionar.AppendLine("/*******************************************************************************************************************************/")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	ISNULL(#tmpTablaResultados.Cod_Usu,'') AS Cod_Usu,")
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Nom_Usu, ")
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Año,")
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Mes,")
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Dia,") 
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Fecha_Entrada,	")		
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Fecha_Salida_Almuerzo,")	
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Fecha_Entrada_Almuerzo,")	
			loComandoSeleccionar.AppendLine("		#tmpTablaResultados.Fecha_Salida,")			
			loComandoSeleccionar.AppendLine("		CASE ")
			loComandoSeleccionar.AppendLine("			WHEN #tmpTablaResultados.Fecha_Entrada  = '1900-01-01 00:00:00.000' THEN 0")
			loComandoSeleccionar.AppendLine("			WHEN #tmpTablaResultados.Fecha_Entrada <> '1900-01-01 00:00:00.000' AND #tmpTablaResultados.Fecha_Salida_Almuerzo = '1900-01-01 00:00:00.000' THEN 0")
			loComandoSeleccionar.AppendLine("			WHEN #tmpTablaResultados.Fecha_Entrada <> '1900-01-01 00:00:00.000' AND #tmpTablaResultados.Fecha_Salida_Almuerzo <> '1900-01-01 00:00:00.000'   ")
			loComandoSeleccionar.AppendLine("					THEN DATEDIFF(ss,#tmpTablaResultados.Fecha_Entrada,#tmpTablaResultados.Fecha_Salida_Almuerzo)  ")
			loComandoSeleccionar.AppendLine("		END AS  Total_Seg_Tra_Man,   ")
			loComandoSeleccionar.AppendLine("		CASE ")
			loComandoSeleccionar.AppendLine("			WHEN #tmpTablaResultados.Fecha_Entrada  = '1900-01-01 00:00:00.000' THEN 0  ")
			loComandoSeleccionar.AppendLine("			WHEN #tmpTablaResultados.Fecha_Entrada <> '1900-01-01 00:00:00.000' AND #tmpTablaResultados.Fecha_Salida_Almuerzo = '1900-01-01 00:00:00.000' THEN 0")
			loComandoSeleccionar.AppendLine("		WHEN #tmpTablaResultados.Fecha_Entrada_Almuerzo = '1900-01-01 00:00:00.000' OR #tmpTablaResultados.Fecha_Salida = '1900-01-01 00:00:00.000' THEN 0 ")
			loComandoSeleccionar.AppendLine("		WHEN #tmpTablaResultados.Fecha_Entrada <> '1900-01-01 00:00:00.000' AND #tmpTablaResultados.Fecha_Entrada_Almuerzo <> '1900-01-01 00:00:00.000' AND	")
			loComandoSeleccionar.AppendLine("				 #tmpTablaResultados.Fecha_Salida <> '1900-01-01 00:00:00.000' THEN DATEDIFF(ss,#tmpTablaResultados.Fecha_Entrada_Almuerzo,#tmpTablaResultados.Fecha_Salida)")
			loComandoSeleccionar.AppendLine("		END AS  Total_Seg_Tra_Tar,")   
			loComandoSeleccionar.AppendLine("		CASE  ")
			loComandoSeleccionar.AppendLine("			WHEN  (#tmpTablaResultados.Fecha_Salida_Almuerzo) <> '1900-01-01 00:00:00.000' AND MAX(#tmpTablaResultados.Fecha_Entrada_Almuerzo) <> '1900-01-01 00:00:00.000'")
			loComandoSeleccionar.AppendLine("				THEN DATEDIFF(ss,#tmpTablaResultados.Fecha_Entrada_Almuerzo,#tmpTablaResultados.Fecha_Salida_Almuerzo) ELSE 0	")
			loComandoSeleccionar.AppendLine("		END AS Total_Seg_Alm ")		
			loComandoSeleccionar.AppendLine("FROM #tmpTablaResultados   ")
			loComandoSeleccionar.AppendLine("WHERE Cod_Usu <> '' ")
			loComandoSeleccionar.AppendLine("GROUP BY Cod_Usu,Nom_Usu,Año,Mes,Dia,Fecha_Entrada,Fecha_Salida_Almuerzo,Fecha_Entrada_Almuerzo,Fecha_Salida")
			loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & ", Año ASC, Mes Desc, Dia Desc  ")
			loComandoSeleccionar.AppendLine("  ")
			loComandoSeleccionar.AppendLine("/*******************************************************************************************************************************/   ")
			loComandoSeleccionar.AppendLine("  ")
			loComandoSeleccionar.AppendLine("  ")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTablaFinal")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTablaTemporal  ")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTablaResultados")
            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_Usuarios_Ipos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_Usuarios.ReportSource = loObjetoReporte

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
' MAT: 29/07/11: Codigo Inicial
'-------------------------------------------------------------------------------------------'
