'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAccesos_Diarios"
'-------------------------------------------------------------------------------------------'
Partial Class rAccesos_Diarios
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try
            Dim ldFecha_Desde As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim ldFecha_Hasta As Date = CDate(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            'Dim loComandoSeleccionar2 As New StringBuilder()
            'Dim loComandoSeleccionar3 As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT     ")
			loComandoSeleccionar.AppendLine("			Accesos.Documento						AS Documento,")
			loComandoSeleccionar.AppendLine("			Accesos.Tipo							AS Tipo,")
			loComandoSeleccionar.AppendLine("			DATEPART(DAYOFYEAR, Accesos.Registro)	AS DiadelAño,")
			loComandoSeleccionar.AppendLine("			DATEPART(YEAR, Accesos.Registro)		AS Año,")
			loComandoSeleccionar.AppendLine("			DATEPART(HOUR, Accesos.Registro)		AS Hora,")
			loComandoSeleccionar.AppendLine("			DATEPART(WEEKDAY, Accesos.Registro)		AS DiaSemana,")
			loComandoSeleccionar.AppendLine("			Accesos.Cod_Tra							AS Cod_Tra,")
			loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra					AS Nom_Tra,")
			loComandoSeleccionar.AppendLine("			Accesos.Registro						AS Registro,")
			loComandoSeleccionar.AppendLine("			Trabajadores.Cod_Tur					AS Cod_Tur,")
			loComandoSeleccionar.AppendLine("			Trabajadores.Mul_Ing					AS Mul_Ing")
			loComandoSeleccionar.AppendLine("INTO		#tmpAccesos")
			loComandoSeleccionar.AppendLine("FROM		Accesos")
			loComandoSeleccionar.AppendLine("			JOIN Trabajadores	ON Trabajadores.Cod_Tra = Accesos.Cod_Tra")
            loComandoSeleccionar.AppendLine("WHERE  	")
            loComandoSeleccionar.AppendLine("         	Accesos.Registro BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("         	AND Accesos.Cod_Tra BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("         	AND Trabajadores.Cod_Dep BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         	AND Trabajadores.Cod_Car BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Trabajadores.Status ='A'")
            loComandoSeleccionar.AppendLine("			AND Accesos.Status <>'Anulado'")
            loComandoSeleccionar.AppendLine("ORDER BY	Accesos.Cod_Tra, Accesos.Registro")

			loComandoSeleccionar.AppendLine("SELECT		#tmpAccesos.Documento				,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Tipo					,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.DiadelAño				,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Año						,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Hora					,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.DiaSemana				,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Cod_Tra					,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Nom_Tra					,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Cod_Tur					,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Registro				,")
			loComandoSeleccionar.AppendLine("			#tmpAccesos.Mul_Ing					AS AlmuerzoFlexible,")
			loComandoSeleccionar.AppendLine("			Turnos.Tip_Tur						,")
			loComandoSeleccionar.AppendLine("			CASE ")
			loComandoSeleccionar.AppendLine("				WHEN (Turnos.tip_tur = 'fijo') ")
			loComandoSeleccionar.AppendLine("					THEN Turnos.hor_ini")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=1)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.dom_ini1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=2)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.lun_ini1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=3)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.mar_ini1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=4)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.mie_ini1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=5)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.jue_ini1")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=6)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.vie_ini1")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=7)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.sab_ini1")
			loComandoSeleccionar.AppendLine("				ELSE ")
			loComandoSeleccionar.AppendLine("					'00:00'")
			loComandoSeleccionar.AppendLine("			END										AS hor_ini,")
			loComandoSeleccionar.AppendLine("			CASE ")
			loComandoSeleccionar.AppendLine("				WHEN (Turnos.tip_tur = 'fijo') ")
			loComandoSeleccionar.AppendLine("					THEN Turnos.com_ini")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=1)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.dom_fin1")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=2)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.lun_fin1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=3)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.mar_fin1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=4)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.mie_fin1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=5)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.jue_fin1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=6)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.vie_fin1")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=7)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.sab_fin1")
			loComandoSeleccionar.AppendLine("				ELSE ")
			loComandoSeleccionar.AppendLine("					'00:00'")
			loComandoSeleccionar.AppendLine("			END										AS com_ini,")	  
			loComandoSeleccionar.AppendLine("			CASE ")
			loComandoSeleccionar.AppendLine("				WHEN (Turnos.tip_tur = 'fijo') ")
			loComandoSeleccionar.AppendLine("					THEN Turnos.com_fin")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=1)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.dom_ini2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=2)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.lun_ini2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=3)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.mar_ini2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=4)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.mie_ini2")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=5)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.jue_ini2")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=6)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.vie_ini2")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=7)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.sab_ini2")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					'00:00'")
            loComandoSeleccionar.AppendLine("			END										AS com_fin,")
            loComandoSeleccionar.AppendLine("			CASE ")
			loComandoSeleccionar.AppendLine("				WHEN (Turnos.tip_tur = 'fijo') ")
			loComandoSeleccionar.AppendLine("					THEN Turnos.hor_fin")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=1)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.dom_fin2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=2)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.lun_fin2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=3)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.mar_fin2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=4)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.mie_fin2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=5)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.jue_fin2")
			loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=6)")
			loComandoSeleccionar.AppendLine("					THEN Turnos.vie_fin2")
            loComandoSeleccionar.AppendLine("				WHEN (#tmpAccesos.DiaSemana=7)")
            loComandoSeleccionar.AppendLine("					THEN Turnos.sab_fin2")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					'00:00'")
            loComandoSeleccionar.AppendLine("			END										AS hor_fin,")
            loComandoSeleccionar.AppendLine("			min_ret									AS MaximoRetardoPermitido,")
            loComandoSeleccionar.AppendLine("			blo_fue									AS BloquearFueraTurno")
            loComandoSeleccionar.AppendLine("FROM		#tmpAccesos")
            loComandoSeleccionar.AppendLine("	JOIN	Turnos ON Turnos.Cod_Tur = #tmpAccesos.Cod_Tur")

															



            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


	'******************************************************************************************
	' Se Procesa manualmente los datos
	'******************************************************************************************

		Dim loTabla As New DataTable("curReportes")
		Dim loColumna As DataColumn 
		
		loColumna = New DataColumn("cod_tra", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Cod_Tur", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Tip_Tur", getType(String))
		loColumna.MaxLength = 30
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("nom_tra", getType(String))
		loColumna.MaxLength = 100
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("fecha", getType(String))
		loColumna.MaxLength = 15
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("registro", getType(Date))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("dia", getType(String))
		loColumna.MaxLength = 10
		loColumna.DefaultValue = "LUN"
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Entrada1", getType(String))
		loColumna.MaxLength = 15
		loColumna.DefaultValue = "[*]"
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Salida1", getType(String))
		loColumna.MaxLength = 15
		loColumna.DefaultValue = "[*]"
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Entrada2", getType(String))
		loColumna.MaxLength = 15
		loColumna.DefaultValue = "[*]"
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Salida2", getType(String))
		loColumna.MaxLength = 15
		loColumna.DefaultValue = "[*]"
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("MinutosAlmuerzo", GetType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("hor_ini", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("com_ini", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)

		loColumna = New DataColumn("com_fin", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)

		loColumna = New DataColumn("hor_fin", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Tiempo_Retrazo", getType(Integer))
		loTabla.Columns.Add(loColumna)

		loColumna = New DataColumn("Cantidad_Retrazo", getType(Integer))
		loTabla.Columns.Add(loColumna)

		loColumna = New DataColumn("AlmuerzoFlexible", getType(Boolean))   'AlmuerzoFlexible,MaximoRetardoPermitido,BloquearFueraTurno 
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("MaximoRetardoPermitido", getType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("BloquearFueraTurno", getType(Boolean))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("NoMarco", getType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Retrazo1", getType(Integer))
		loTabla.Columns.Add(loColumna)

		loColumna = New DataColumn("Retrazo2", getType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("NoVino", getType(Integer))
		loTabla.Columns.Add(loColumna)
		
		
	'-------------------------------------------------------------------------------------------'
	' Para cada fila de los datos originales, genera una nueva fila para mostrar en pantalla.	'
	'-------------------------------------------------------------------------------------------'
		Dim loNuevaFila As DataRow
		Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
		For lnNumeroFila As Integer = 0 To lnTotalFilas - 1   
			
			Dim loFila As DataRow = laDatosReporte.Tables(0).Rows(lnNumeroFila)
			loNuevaFila = loTabla.NewRow()
			loTabla.Rows.Add(loNuevaFila)
			
			Dim lcCod_Tra	As String = Trim(loFila("cod_tra"))
			Dim lcRegistro	As String = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")
			
			loNuevaFila.Item("cod_tra")		= lcCod_Tra
			loNuevaFila.Item("nom_tra")		= Trim(loFila("nom_tra"))
			loNuevaFila.Item("registro")	= loFila("registro")
			loNuevaFila.Item("fecha")		= lcRegistro
			loNuevaFila.Item("hor_ini")		= loFila("hor_ini")
			loNuevaFila.Item("com_ini")		= loFila("com_ini")
			loNuevaFila.Item("com_fin")		= loFila("com_fin")
			loNuevaFila.Item("hor_fin")		= loFila("hor_fin")
			loNuevaFila.Item("Cod_Tur")		= loFila("Cod_Tur")
			loNuevaFila.Item("Tip_Tur")		= loFila("Tip_Tur")	

			loNuevaFila.Item("AlmuerzoFlexible")		= loFila("AlmuerzoFlexible")
			loNuevaFila.Item("MaximoRetardoPermitido")	= loFila("MaximoRetardoPermitido")
			loNuevaFila.Item("BloquearFueraTurno")		= loFila("BloquearFueraTurno")
						
			If (Trim(loFila("Tipo"))		= "Entrada") Then 

				loNuevaFila.Item("Entrada1") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
				
				lnNumeroFila += 1
				
				If	(lnNumeroFila < lnTotalFilas) Then 
					
					loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
					
					If	(Trim(loFila("Tipo"))		= "Salida")		And _
						(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
						(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then
						
						loNuevaFila.Item("Salida1") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
					
						lnNumeroFila += 1
						
						If	(lnNumeroFila < lnTotalFilas) Then 
							loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
							
							If	(Trim(loFila("Tipo"))		= "Entrada")	And _
								(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
								(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then
								
								loNuevaFila.Item("Entrada2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")

								lnNumeroFila += 1
								
								If	(lnNumeroFila < lnTotalFilas) Then 
									
									loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
									
									If	(Trim(loFila("Tipo"))		= "Salida")		And _
										(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
										(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then

										loNuevaFila.Item("Salida2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
										
									Else
										
										loNuevaFila.Item("Salida2") = "[*]"
										lnNumeroFila -= 1
										
									End If
									
								End If
										
							Else ' Hay entrada 1 y salida 1
							
								loNuevaFila.Item("Entrada2") = "[*]"

								If	(Trim(loFila("Tipo"))		= "Salida")		And _
									(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
									(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then

									loNuevaFila.Item("Salida2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
									
								Else
									
									loNuevaFila.Item("Salida2") = "[*]"
									lnNumeroFila -= 1
									
								End If
									
							End If
							
						End If

						
					Else ' Primera Entrada Sin Salida

						loNuevaFila.Item("Salida1") = "[*]"

							If	(Trim(loFila("Tipo"))		= "Entrada")	And _
								(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
								(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then
								
								loNuevaFila.Item("Entrada2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")

								lnNumeroFila += 1
								
								If	(lnNumeroFila < lnTotalFilas) Then 
									
									loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
									
									If	(Trim(loFila("Tipo"))		= "Salida")		And _
										(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
										(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then

										loNuevaFila.Item("Salida2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
										
									Else
										
										loNuevaFila.Item("Salida2") = "[*]"
										lnNumeroFila -= 1
										
									End If
									
								End If
										
							Else
							
								loNuevaFila.Item("Entrada2") = "[*]"
								loNuevaFila.Item("Salida2")	= "[*]"
								lnNumeroFila -= 1
									
							End If
							
					End If
					
				End If
				
			Else ' Lo primero es una salida
				
				loNuevaFila.Item("Entrada1")	= "[*]"
				loNuevaFila.Item("Salida1")		= Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")
				
				lnNumeroFila += 1
				
				If	(lnNumeroFila < lnTotalFilas) Then 
					
					loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
					
					If	(Trim(loFila("Tipo"))		= "Entrada")	And _
						(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
						(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then
						
						loNuevaFila.Item("Entrada2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")

						lnNumeroFila += 1
						
						If	(lnNumeroFila < lnTotalFilas) Then 
							
							loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
							
							If	(Trim(loFila("Tipo"))		= "Salida")		And _
								(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
								(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then

								loNuevaFila.Item("Salida2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")

							Else
								
								loNuevaFila.Item("Salida2") = "[*]"
								lnNumeroFila -= 1
								
							End If
							
						End If
								
					Else
					
						loNuevaFila.Item("Entrada2") = "[*]"

						If	(Trim(loFila("Tipo"))		= "Salida")		And _
							(Trim(loFila("cod_tra"))	= lcCod_Tra)	And _
							(Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "dd/MM/yyyy")	= lcRegistro)	Then

							loNuevaFila.Item("Salida2") = Microsoft.VisualBasic.Format(CDate(loFila("Registro")), "hh:mm:ss tt")

						Else
							
							loNuevaFila.Item("Salida2") = "[*]"
							lnNumeroFila -= 1
							
						End If
							
					End If
					
				End If				
				
			End If
			
			loTabla.AcceptChanges()
		
		Next lnNumeroFila
		
'AlmuerzoFlexible, MaximoRetardoPermitido, BloquearFueraTurno
		
	'---------------------------------------------------------------------------------------'
	' Calcula los almuerzos y retrazos, y coloca las faltas (días que no marcó).			'
	'---------------------------------------------------------------------------------------'
		For Each loFila As DataRow In loTabla.Rows
			
			Dim lnMinutosAlmuerzo	As Integer	= 0  
			Dim ldInicio_Almuerzo	As Date
			Dim ldFin_Almuerzo		As Date
			Dim ldHor_Ini			As Date		= CDate(loFila("hor_ini"))
			Dim ldCom_Ini			As Date		= CDate(loFila("com_ini"))
			Dim ldCom_Fin			As Date		= CDate(loFila("com_fin"))
			Dim ldHor_Fin			As Date		= CDate(loFila("hor_fin"))
			Dim lnTiempo_Retrazo	As Integer	= 0
			Dim lnCantidad_Retrazo	As Integer	= 0
			Dim lnNoMarco			As Integer	= 0

		'---------------------------------------------------------------------------------------'
		' Coloca en la segunda parte del turno los accesos colocados en el primero por error	'
		'---------------------------------------------------------------------------------------'
			'If	(Trim(loFila("cod_tra")) = "wjr")			Then
				'loFila("NoVino") = 1
			'End If
		'---------------------------------------------------------------------------------------'
		' Coloca en la segunda parte del turno los accesos colocados en el primero por error	'
		'---------------------------------------------------------------------------------------'
			If	(Trim(loFila("Entrada1")) <> "[*]")			AndAlso _
				(Trim(loFila("Entrada2")) = "[*]")			AndAlso _
				(CDate(loFila("Entrada1")) > ldCom_Ini)		Then 
				
				loFila("Entrada2") = loFila("Entrada1")
				loFila("Entrada1") = "[*]"
				
			End If
			
			If	(Trim(loFila("Salida1")) <> "[*]")			AndAlso _
				(Trim(loFila("Salida2")) = "[*]")			Then 
				
				If	(Trim(loFila("Entrada2")) <> "[*]")						AndAlso _
					(CDate(loFila("Salida1")) > CDate(loFila("Entrada2")))	Then 
				
					loFila("Salida2") = loFila("Salida1")
					loFila("Salida1") = "[*]"
					
				ElseIf (Trim(loFila("Entrada2")) = "[*]")		AndAlso _
					(CDate(loFila("Salida1")) > ldCom_Ini)	Then 
				
					loFila("Salida2") = loFila("Salida1")
					loFila("Salida1") = "[*]"
					
				End If
				
			End If
			
			
			'If	(Trim(loFila("Salida1")) <> "[*]")			AndAlso _
			'	(Trim(loFila("Salida2")) = "[*]")			AndAlso _
			'	(CDate(loFila("Salida1")) > ldCom_Ini)		Then 
				
			'	loFila("Salida2") = loFila("Salida1")
			'	loFila("Salida1") = "[*]"
				
			'End If
			
			
		'---------------------------------------------------------------------------------------'
		' Calcula el tiempo de descanso (almuerzo) en minutos.									'
		'---------------------------------------------------------------------------------------'
			If	(Trim(loFila("Salida1")) <> "[*]")	AndAlso _
				(loFila("Entrada2") <> "[*]")
				
				ldInicio_Almuerzo	= CDate(loFila("Salida1")) 
				ldFin_Almuerzo		= CDate(loFila("Entrada2")) 
				
				lnMinutosAlmuerzo	= (ldFin_Almuerzo - ldInicio_Almuerzo).TotalMinutes 
				
			End If
			
			loFila("MinutosAlmuerzo") = lnMinutosAlmuerzo
			
		'---------------------------------------------------------------------------------------'
		' Calcula el tiempo de retrazo, el número de retrazos y el numero de veces que no marcó.'
		'---------------------------------------------------------------------------------------'
			
			Dim lnMinutosRetrasoEntrada1 As Integer 
			If (Trim(loFila("Entrada1")) <> "[*]") Then 
			
				lnMinutosRetrasoEntrada1 = (CDate(loFila("Entrada1")) - ldHor_Ini).TotalMinutes() - CInt(loFila("MaximoRetardoPermitido"))
			
				If	(lnMinutosRetrasoEntrada1 > 0D)				Then 
					
					lnTiempo_Retrazo	+= lnMinutosRetrasoEntrada1 + CInt(loFila("MaximoRetardoPermitido"))
					lnCantidad_Retrazo	+= 1
					loFila("Retrazo1")	 = 1
					
				End If

			End If
			
			Dim lnMinutosRetrasoEntrada2 As Integer 
			If (Trim(loFila("Entrada2")) <> "[*]") Then 
			
				If CBool(loFila("AlmuerzoFlexible") AndAlso Trim(loFila("Salida1")) <> "[*]") Then
					lnMinutosRetrasoEntrada2 = (CDate(loFila("Entrada2")) - CDate(loFila("Salida1"))).TotalMinutes() 'Tiempo almuerzo usado
					lnMinutosRetrasoEntrada2 = lnMinutosRetrasoEntrada2 - (ldCom_Fin - ldCom_Ini).TotalMinutes() - CInt(loFila("MaximoRetardoPermitido"))
				Else 
					lnMinutosRetrasoEntrada2 = (CDate(loFila("Entrada2")) - ldCom_Fin).TotalMinutes() - CInt(loFila("MaximoRetardoPermitido"))
				End If
			
				If	(lnMinutosRetrasoEntrada2 > 0D)				Then 
				
				'---------------------------------------------------------------------------------------'
				' Si se pasó del tiempo máximo de retrazo, entonces los minutos de retrazo se cuentan completos.
				'---------------------------------------------------------------------------------------'
					lnTiempo_Retrazo	+= lnMinutosRetrasoEntrada2 + CInt(loFila("MaximoRetardoPermitido"))
					lnCantidad_Retrazo	+= 1
					loFila("Retrazo2")	 = 1
					
				End If

			End If

			'If	(Trim(loFila("Entrada2")) <> "[*]")			AndAlso _
			'	(CDate(loFila("Entrada2")) > (ldCom_Fin) )	Then 
				
			'	lnTiempo_Retrazo	+= (CDate(loFila("Entrada2")) - (ldCom_Fin)).TotalMinutes()
			'	lnCantidad_Retrazo	+= 1
			'	loFila("Retrazo2")	 = 1
				
			'End If

			If	(Trim(loFila("Entrada1")) = "[*]")				Then 
				
				loFila("Entrada1") = "[No Marcó]"
				lnNoMarco += 1
				
			End If
			
			If	(Trim(loFila("Salida1")) = "[*]")				Then 
				
				loFila("Salida1") = "[No Marcó]"
				lnNoMarco += 1
				
			End If
			
			If	(Trim(loFila("Entrada2")) = "[*]")				Then 
				
				loFila("Entrada2") = "[No Marcó]"
				lnNoMarco += 1
				
			End If
			
			If	(Trim(loFila("Salida2")) = "[*]")				Then 
				
				loFila("Salida2") = "[No Marcó]"
				lnNoMarco += 1
				
			End If
			
			
			Dim lnDiaSemana As System.DayOfWeek = DirectCast(loFila("registro"), Date).DayOfWeek 
			
			Select Case lnDiaSemana
			
				Case DayOfWeek.Monday
					loFila("dia") = "LUN"
					
				Case DayOfWeek.Tuesday
					loFila("dia") = "MAR"
					
				Case DayOfWeek.Wednesday
					loFila("dia") = "MIE"
					
				Case DayOfWeek.Thursday
					loFila("dia") = "JUE"
					
				Case DayOfWeek.Friday
					loFila("dia") = "VIE"
					
				Case DayOfWeek.Saturday
					loFila("dia") = "SAB"
					
				Case DayOfWeek.Sunday
					loFila("dia") = "DOM"
					
			End Select
						
			loFila("Cantidad_Retrazo")	=   lnCantidad_Retrazo
			loFila("Tiempo_Retrazo")	=   lnTiempo_Retrazo
			loFila("NoMarco")			=   lnNoMarco
					
		Next loFila
		
		loTabla.AcceptChanges()


	'---------------------------------------------------------------------------------------'
	' Busca los usuarios que no estén incluidos	en el SELECT anterior(los que no han venido)'
	'---------------------------------------------------------------------------------------'
        loComandoSeleccionar.Length = 0
        
        loComandoSeleccionar.AppendLine("SELECT			''										AS Documento,")
        loComandoSeleccionar.AppendLine("				''										AS Tipo,")
		loComandoSeleccionar.AppendLine("				''										AS DiadelAño,")
		loComandoSeleccionar.AppendLine("				''										AS Año,")
		loComandoSeleccionar.AppendLine("				''										AS Hora,")
		loComandoSeleccionar.AppendLine("				''										AS DiaSemana,")
		loComandoSeleccionar.AppendLine("				Trabajadores.Cod_Tra					AS Cod_Tra,")
        loComandoSeleccionar.AppendLine("				Trabajadores.Nom_Tra					AS Nom_Tra,")
        loComandoSeleccionar.AppendLine("				GETDATE()								AS Registro,")
        loComandoSeleccionar.AppendLine("				Trabajadores.Cod_Tur					AS Cod_Tur,")
        loComandoSeleccionar.AppendLine("				Trabajadores.Mul_Ing					AS AlmuerzoFlexible,")
        loComandoSeleccionar.AppendLine("				''										AS Tip_Tur,")
        loComandoSeleccionar.AppendLine("				'00:00'									AS hor_ini,")
        loComandoSeleccionar.AppendLine("				'00:00'									AS com_ini,")
        loComandoSeleccionar.AppendLine("				'00:00'									AS com_fin,")
        loComandoSeleccionar.AppendLine("				'00:00'									AS hor_fin,")
        loComandoSeleccionar.AppendLine("				CAST(0 AS INT)							AS MaximoRetardoPermitido,")
        loComandoSeleccionar.AppendLine("				CAST(0 AS BIT)							AS BloquearFueraTurno")
        loComandoSeleccionar.AppendLine("FROM			Trabajadores")
        loComandoSeleccionar.AppendLine("WHERE			Trabajadores.Cod_Tra BETWEEN " & lcParametro1Desde)
        loComandoSeleccionar.AppendLine("         			AND " & lcParametro1Hasta)
        loComandoSeleccionar.AppendLine("         	AND Trabajadores.Cod_Dep BETWEEN " & lcParametro2Desde)
        loComandoSeleccionar.AppendLine("         			AND " & lcParametro2Hasta)
        loComandoSeleccionar.AppendLine("         	AND Trabajadores.Cod_Car BETWEEN " & lcParametro3Desde)
        loComandoSeleccionar.AppendLine("         			AND " & lcParametro3Hasta) 		
		loComandoSeleccionar.AppendLine("			AND Trabajadores.Status ='A'")
		loComandoSeleccionar.AppendLine("ORDER BY	Trabajadores.Cod_Tra")

		Dim TablaRelleno As DataTable 
			
		TablaRelleno = (New goDatos()).mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "").Tables(0)
		
	'---------------------------------------------------------------------------------------'
	' Rellena la tabla con los trabajadores/días faltantes.									'
	'---------------------------------------------------------------------------------------'
		Dim ldRango As TimeSpan = (ldFecha_Hasta - ldFecha_Desde)
		Dim ldFecha As DateTime
		Dim lcInicioDia	As String
		Dim lcFinDia	As String	
		Dim lcSeleccion As String 
		
		For Each loRenglonRelleno As DataRow In TablaRelleno.Rows
			
			Dim lcTrabajador As String = goServicios.mObtenerCampoFormatoSQL(loRenglonRelleno("cod_tra"))
			ldFecha = ldFecha_Desde
			
			For lnDia As Integer = 0 To CInt(Math.Truncate(ldRango.TotalDays()))
				
				Dim lfFechaTmp As DateTime = ldFecha.AddDays(lnDia)
				
				lcInicioDia	= "CONVERT('" & CStr(lfFechaTmp.Month()) & "-" & CStr(lfFechaTmp.Day()) & "-" & CStr(lfFechaTmp.Year()) & "', 'System.DateTime')"
				lfFechaTmp = lfFechaTmp.AddDays(1)
				lcFinDia	= "CONVERT('" & CStr(lfFechaTmp.Month()) & "-" & CStr(lfFechaTmp.Day()) & "-" & CStr(lfFechaTmp.Year()) & "', 'System.DateTime')"
				
				lcSeleccion = "Cod_Tra=" & lcTrabajador & " AND registro>" & lcInicioDia & " AND registro<" & lcFinDia
				lfFechaTmp = lfFechaTmp.AddDays(-1)
				
				If loTabla.Select(lcSeleccion).Length <= 0 Then 
					Dim loNuevaFilaRelleno As DataRow = loTabla.NewRow()
					
					loNuevaFilaRelleno("cod_tra") 			= loRenglonRelleno("cod_tra")
					loNuevaFilaRelleno("nom_tra") 			= loRenglonRelleno("nom_tra")
					loNuevaFilaRelleno("fecha")				= Microsoft.VisualBasic.Format(lfFechaTmp, "dd/MM/yyyy")
					loNuevaFilaRelleno("registro")			= lfFechaTmp
					loNuevaFilaRelleno("Entrada1") 			= "[No Vino]"
					loNuevaFilaRelleno("Salida1") 			= "[No Vino]"
					loNuevaFilaRelleno("Entrada2") 			= "[No Vino]"
					loNuevaFilaRelleno("Salida2") 			= "[No Vino]"
					loNuevaFilaRelleno("MinutosAlmuerzo") 	= 0I
					loNuevaFilaRelleno("NoMarco") 			= 0I
					loNuevaFilaRelleno("NoVino") 			= 1I
					loNuevaFilaRelleno("Cantidad_Retrazo")	= 0I
					loNuevaFilaRelleno("Tiempo_Retrazo")	= 0D
					
					Dim lnDiaSemana As System.DayOfWeek = lfFechaTmp.DayOfWeek 
					
					Select Case lnDiaSemana
					
						Case DayOfWeek.Monday
							loNuevaFilaRelleno("dia") = "LUN"
							
						Case DayOfWeek.Tuesday
							loNuevaFilaRelleno("dia") = "MAR"
							
						Case DayOfWeek.Wednesday
							loNuevaFilaRelleno("dia") = "MIE"
							
						Case DayOfWeek.Thursday
							loNuevaFilaRelleno("dia") = "JUE"
							
						Case DayOfWeek.Friday
							loNuevaFilaRelleno("dia") = "VIE"
							
						Case DayOfWeek.Saturday
							loNuevaFilaRelleno("dia") = "SAB"
							
						Case DayOfWeek.Sunday
							loNuevaFilaRelleno("dia") = "DOM"
							
					End Select

					loTabla.Rows.Add(loNuevaFilaRelleno)
					
				End If
				
			Next lnDia
			
		Next loRenglonRelleno
		
		
		'Dim loTablaFinal As New DataTable 
		
		
		'For Each loColumnaFinal As DataColumn In loTabla.Columns
			
		'	loTabla.Columns.Remove(loColumnaFinal)
		'	loTablaFinal.Columns.Add(loColumnaFinal)
			
		'Next loColumnaFinal
		
		'For Each loRenglonFinal As DataRow In loTabla.Select("1=1", "Cod_Tra, Registro")
			
		'	loTablaFinal.ImportRow(loRenglonFinal)
			
		'Next loRenglonFinal
		
		'loTabla.Clear()
		'loTabla = Nothing 
		
		Dim loDatosReporteFinal As New DataSet("curReportes") 
		loDatosReporteFinal.Tables.Add(loTabla)



	'---------------------------------------------------------------------------------------'
	' Se llena el reporte con la tabla nueva												'
	'---------------------------------------------------------------------------------------'
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAccesos_Diarios", loDatosReporteFinal)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrAccesos_Diarios.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception

            'Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
            '              "No se pudo Completar el Proceso: " & loExcepcion.Message, _
            '               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
            '               "auto", _
            '               "auto")

        'End Try
        
        
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
' CMS: 02/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 02/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 27/01/10: Modificado para agregar procesamiento de "Minutos permitidos de retrazo" y	'
'				 "Permitir almuerzo flexible". Eliminado el resaltado gris de los renglones.'
'-------------------------------------------------------------------------------------------'
