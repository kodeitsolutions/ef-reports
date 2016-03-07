'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCobros_Anualmente"
'-------------------------------------------------------------------------------------------'
Partial Class grCobros_Anualmente
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Cobros.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine(" 		    WHEN Detalles_Cobros.Tip_Ope = 'Efectivo' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Efectivo,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Ticket' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Ticket,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Cheque' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Cheque,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Deposito' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Deposito,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Transferencia")
            loComandoSeleccionar.AppendLine(" INTO #temp_cobros")
            loComandoSeleccionar.AppendLine(" FROM Cobros Cobros")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Cobros.Fec_ini BETWEEN DATEADD (YYYY , -10, " & lcParametro0Hasta & " )")
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Cobros.fec_ini), Detalles_Cobros.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ORDER BY DATEPART(YEAR, Cobros.fec_ini) ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	Año,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Deposito) AS Deposito, ")
            loComandoSeleccionar.AppendLine("    	SUM(Transferencia) AS Transferencia ")
            loComandoSeleccionar.AppendLine(" INTO #temp_totcobros  ")
            loComandoSeleccionar.AppendLine(" FROM #temp_cobros  ")
            loComandoSeleccionar.AppendLine(" GROUP BY Año")
            loComandoSeleccionar.AppendLine(" ORDER BY Año ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            'loComandoSeleccionar.AppendLine("       #temp_cobros.Año AS Año,    ")
            loComandoSeleccionar.AppendLine("       #temp_totcobros.Año AS Año,    ")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Efectivo) AS Efectivo,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Ticket) AS Ticket,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Cheque)   AS Cheque,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Tarjeta)  AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Deposito) AS Deposito,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine("       SUM(#temp_totcobros.Efectivo + #temp_totcobros.Cheque + #temp_totcobros.Tarjeta + #temp_totcobros.Deposito + #temp_totcobros.Transferencia + #temp_totcobros.Ticket) AS Total_Cobros")
            loComandoSeleccionar.AppendLine(" Into #Final   ")
            'loComandoSeleccionar.AppendLine(" FROM	#temp_cobros   ")
            'loComandoSeleccionar.AppendLine(" FULL JOIN  #temp_totcobros ON (#temp_cobros.Año = #temp_totcobros.Año) ")
            'loComandoSeleccionar.AppendLine(" GROUP BY #temp_cobros.Año")
            
            loComandoSeleccionar.AppendLine(" FROM	#temp_totcobros    ")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_totcobros .Año")
            
            loComandoSeleccionar.AppendLine(" Union All ")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -10, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -9, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -8, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -7, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -6, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -5, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -4, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -3, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -2, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, DATEADD(yyyy , -1, " & lcParametro0Hasta & " )) As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            loComandoSeleccionar.AppendLine(" Union All")
            loComandoSeleccionar.AppendLine(" Select Datepart(yyyy, " & lcParametro0Hasta & ")  As Año, 0 As Efectivo, 0 As Ticket, 0 As Cheque, 0 As Tarjeta, 0 As Deposito, 0 As Transferencia, 0 As Total_Cobros")
            'loComandoSeleccionar.AppendLine(" ORDER BY #temp_cobros.Año ")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_totcobros.Año ")
            
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		Año,     ")
            loComandoSeleccionar.AppendLine(" 		SUM(Efectivo) AS Efectivo, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Ticket) AS Ticket, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Cheque)   AS Cheque, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Tarjeta)  AS Tarjeta, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Deposito) AS Depósito, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Transferencia) AS Transferencia, ")
            loComandoSeleccionar.AppendLine(" 		SUM(Total_Cobros) AS Total_Cobros ")
            loComandoSeleccionar.AppendLine(" FROM #Final ")
			loComandoSeleccionar.AppendLine(" WHERE Año BETWEEN")
			loComandoSeleccionar.AppendLine(" (Case ")
			loComandoSeleccionar.AppendLine(" 	When  " & lcParametro0Desde & " > DATEADD(YYYY , -10, " & lcParametro0Hasta & ") Then Datepart(yyyy,DATEADD(YYYY , -10, " & lcParametro0Hasta & "))")
			loComandoSeleccionar.AppendLine(" 	Else Datepart(yyyy," & lcParametro0Desde & ")")
			loComandoSeleccionar.AppendLine(" 	End )")
			loComandoSeleccionar.AppendLine("   AND Datepart(yyyy,'20100611 23:59:59.998')")
            
            loComandoSeleccionar.AppendLine(" GROUP BY Año ")
            loComandoSeleccionar.AppendLine(" ORDER BY Año  ")
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

	''******************************************************************************************
	'' Inicio Se Procesa manualmetne los datos
	''******************************************************************************************

	'	'Tabla con las listas desplegables
	'	Dim loTabla As New DataTable("curReportes")
	'	Dim loColumna As DataColumn 
		
	'	loColumna = New DataColumn("Año", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Semana", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Efectivo", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Ticket", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Cheque", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Tarjeta", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Deposito", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Transferencia", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		
	'	loColumna = New DataColumn("Total_Cobros", getType(Integer))
	'	loTabla.Columns.Add(loColumna)
		

	'	Dim loNuevaFila As DataRow
	'	Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
			
		
	'	For lnNumeroFila As Integer = 0 To lnTotalFilas - 1  
		
	'		Dim loFila As DataRow = laDatosReporte.Tables(0).Rows(lnNumeroFila)
	'		'loNuevaFila = loTabla.NewRow()
	'		'loTabla.Rows.Add(loNuevaFila)
		
	'		if lnNumeroFila = 0 Then 

	'			loNuevaFila = loTabla.NewRow()
	'			loTabla.Rows.Add(loNuevaFila)
				
	'			loNuevaFila.Item("Año")					= loFila("Año")
	'			loNuevaFila.Item("Semana")				= loFila("Semana")
	'			loNuevaFila.Item("Efectivo")			= loFila("Efectivo")
	'			loNuevaFila.Item("Ticket")				= loFila("Ticket")
	'			loNuevaFila.Item("Cheque")				= loFila("Cheque")
	'			loNuevaFila.Item("Tarjeta")				= loFila("Tarjeta")		
	'			loNuevaFila.Item("Deposito")			= loFila("Deposito")		
	'			loNuevaFila.Item("Transferencia")		= loFila("Transferencia")		
	'			loNuevaFila.Item("Total_Cobros")		= loFila("Total_Cobros")	
				
	'			loTabla.AcceptChanges()

	'		End If

	'		if lnNumeroFila > 0 Then 			
	'			If (loFila("Semana") = 53 ) Then 

	'				If (laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana") = 1) Then 
						
	'					loNuevaFila = loTabla.NewRow()
	'					loTabla.Rows.Add(loNuevaFila)
						
	'					loNuevaFila.Item("Año")					= loFila("Año")
	'					loNuevaFila.Item("Semana")				= loFila("Semana")
	'					loNuevaFila.Item("Efectivo")			= loFila("Efectivo")
	'					loNuevaFila.Item("Ticket")				= loFila("Ticket")
	'					loNuevaFila.Item("Cheque")				= loFila("Cheque")
	'					loNuevaFila.Item("Tarjeta")				= loFila("Tarjeta")		
	'					loNuevaFila.Item("Deposito")			= loFila("Deposito")		
	'					loNuevaFila.Item("Transferencia")		= loFila("Transferencia")		
	'					loNuevaFila.Item("Total_Cobros")		= loFila("Total_Cobros")	
						
	'					loTabla.AcceptChanges()
					
	'				ELse
					
	'					For i As  Integer = 1 To (laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana") - 1) 
						
	'						loNuevaFila = loTabla.NewRow()
	'						loTabla.Rows.Add(loNuevaFila)
							
	'						loNuevaFila.Item("Año")					= laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Año")
	'						loNuevaFila.Item("Semana")				= (laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana") - i)
	'						loNuevaFila.Item("Efectivo")			= 0
	'						loNuevaFila.Item("Ticket")				= 0
	'						loNuevaFila.Item("Cheque")				= 0
	'						loNuevaFila.Item("Tarjeta")				= 0
	'						loNuevaFila.Item("Deposito")			= 0
	'						loNuevaFila.Item("Transferencia")		= 0
	'						loNuevaFila.Item("Total_Cobros")		= 0
							
	'						loTabla.AcceptChanges()

	'					Next
						
	'					loNuevaFila = loTabla.NewRow()
	'					loTabla.Rows.Add(loNuevaFila)
						
	'					loNuevaFila.Item("Año")					= loFila("Año")
	'					loNuevaFila.Item("Semana")				= loFila("Semana")
	'					loNuevaFila.Item("Efectivo")			= loFila("Efectivo")
	'					loNuevaFila.Item("Ticket")				= loFila("Ticket")
	'					loNuevaFila.Item("Cheque")				= loFila("Cheque")
	'					loNuevaFila.Item("Tarjeta")				= loFila("Tarjeta")		
	'					loNuevaFila.Item("Deposito")			= loFila("Deposito")		
	'					loNuevaFila.Item("Transferencia")		= loFila("Transferencia")		
	'					loNuevaFila.Item("Total_Cobros")		= loFila("Total_Cobros")
						
	'					loTabla.AcceptChanges()
					
	'				End If
					
	'			Else 
				
	'				If (loFila("Semana")+ 1 = laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana")) Then 
					
	'					loNuevaFila = loTabla.NewRow()
	'					loTabla.Rows.Add(loNuevaFila)
						
	'					loNuevaFila.Item("Año")					= loFila("Año")
	'					loNuevaFila.Item("Semana")				= loFila("Semana")
	'					loNuevaFila.Item("Efectivo")			= loFila("Efectivo")
	'					loNuevaFila.Item("Ticket")				= loFila("Ticket")
	'					loNuevaFila.Item("Cheque")				= loFila("Cheque")
	'					loNuevaFila.Item("Tarjeta")				= loFila("Tarjeta")		
	'					loNuevaFila.Item("Deposito")			= loFila("Deposito")		
	'					loNuevaFila.Item("Transferencia")		= loFila("Transferencia")		
	'					loNuevaFila.Item("Total_Cobros")		= loFila("Total_Cobros")	
						
	'					loTabla.AcceptChanges()
					
	'				ELse
					
	'					For i As  Integer = 1 To (laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana") - loFila("Semana")) - 1
						
	'						loNuevaFila = loTabla.NewRow()
	'						loTabla.Rows.Add(loNuevaFila)
							
	'						loNuevaFila.Item("Año")					= laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Año")
	'						loNuevaFila.Item("Semana")				= (laDatosReporte.Tables(0).Rows(lnNumeroFila-1).item("Semana") - i)
	'						loNuevaFila.Item("Efectivo")			= 0
	'						loNuevaFila.Item("Ticket")				= 0
	'						loNuevaFila.Item("Cheque")				= 0
	'						loNuevaFila.Item("Tarjeta")				= 0
	'						loNuevaFila.Item("Deposito")			= 0
	'						loNuevaFila.Item("Transferencia")		= 0
	'						loNuevaFila.Item("Total_Cobros")		= 0
							
	'						loTabla.AcceptChanges()

	'					Next
						
	'					loNuevaFila = loTabla.NewRow()
	'					loTabla.Rows.Add(loNuevaFila)
						
	'					loNuevaFila.Item("Año")					= loFila("Año")
	'					loNuevaFila.Item("Semana")				= loFila("Semana")
	'					loNuevaFila.Item("Efectivo")			= loFila("Efectivo")
	'					loNuevaFila.Item("Ticket")				= loFila("Ticket")
	'					loNuevaFila.Item("Cheque")				= loFila("Cheque")
	'					loNuevaFila.Item("Tarjeta")				= loFila("Tarjeta")		
	'					loNuevaFila.Item("Deposito")			= loFila("Deposito")		
	'					loNuevaFila.Item("Transferencia")		= loFila("Transferencia")		
	'					loNuevaFila.Item("Total_Cobros")		= loFila("Total_Cobros")
						
	'					loTabla.AcceptChanges()
					
	'				End If
				
	'			End If
			
	'		End If
			
	'		'loTabla.AcceptChanges()
			
	'	Next lnNumeroFila
		


	''******************************************************************************************
	'' Fin Se Procesa manualmetne los datos
	''******************************************************************************************

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
            
            Dim Total As Decimal = 0
            
            For Each loFilas As DataRow In laDatosReporte.Tables(0).Rows

                Total = Total + loFilas.Item("Total_Cobros")

            Next loFilas
            
            If Total = 0 And laDatosReporte.Tables(0).Rows.Count >= 0 Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If
            
			'dim loTablaRellenada As New DataSet()
			'loTablaRellenada.Tables.Add(loTabla)
			
            'loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_Anualmente", loTablaRellenada)
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_Anualmente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCobros_Anualmente.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.message, _
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
' CMS: 29/05/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
