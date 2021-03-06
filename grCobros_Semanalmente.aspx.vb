﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCobros_Semanalmente"
'-------------------------------------------------------------------------------------------'
Partial Class grCobros_Semanalmente
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
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
            loComandoSeleccionar.AppendLine("       DATEPART(ww, Cobros.fec_ini) AS Semana,")
            loComandoSeleccionar.AppendLine("       DATEPART(yy, Cobros.fec_ini) AS Año,")
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
            loComandoSeleccionar.AppendLine("       END AS Depósito,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN SUM(ISNULL(Detalles_Cobros.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("           ELSE 0")
            loComandoSeleccionar.AppendLine("       END AS Transferencia")
            loComandoSeleccionar.AppendLine(" INTO #temp_cobros")
            loComandoSeleccionar.AppendLine(" FROM Cobros Cobros")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" WHERE Cobros.Fec_ini BETWEEN DATEADD (ww , -46, " & lcParametro0Hasta & " )")
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
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Cobros.fec_ini), DATEPART(ww, Cobros.fec_ini), Detalles_Cobros.Tip_Ope")
            loComandoSeleccionar.AppendLine(" ORDER BY DATEPART(YEAR, Cobros.fec_ini) DESC, DATEPART(ww, Cobros.fec_ini) DESC")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("    	Año,  ")
            loComandoSeleccionar.AppendLine("    	Semana,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Efectivo) AS Efectivo,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Ticket) AS Ticket,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Cheque) AS Cheque,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Tarjeta) AS Tarjeta,  ")
            loComandoSeleccionar.AppendLine("    	SUM(Depósito) AS Depósito, ")
            loComandoSeleccionar.AppendLine("    	SUM(Transferencia) AS Transferencia ")
            loComandoSeleccionar.AppendLine(" INTO #temp_totcobros  ")
            loComandoSeleccionar.AppendLine(" FROM #temp_cobros  ")
            loComandoSeleccionar.AppendLine(" GROUP BY Año, Semana  ")
            loComandoSeleccionar.AppendLine(" ORDER BY Año DESC, Semana DESC ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine("       #temp_totcobros.Año AS Año,    ")
            loComandoSeleccionar.AppendLine("       #temp_totcobros.Semana AS Semana,    ")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Efectivo) AS Efectivo,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Ticket) AS Ticket,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Cheque)   AS Cheque,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Tarjeta)  AS Tarjeta,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Depósito) AS Depósito,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Transferencia) AS Transferencia,")
            loComandoSeleccionar.AppendLine("       (#temp_totcobros.Efectivo + #temp_totcobros.Cheque + #temp_totcobros.Tarjeta + #temp_totcobros.Depósito + #temp_totcobros.Transferencia + #temp_totcobros.Ticket) AS Total_Cobros,")
            loComandoSeleccionar.AppendLine("		DATEPART(ww, Cast(#temp_totcobros.Año As varchar(4)) + '1231') AS UltimaSemana")
            loComandoSeleccionar.AppendLine(" INTO	#Final")
            loComandoSeleccionar.AppendLine(" FROM	#temp_totcobros   ")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_totcobros.Año,  #temp_totcobros.Semana Asc")
            
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Año,")
            loComandoSeleccionar.AppendLine(" 		Semana,")
            loComandoSeleccionar.AppendLine(" 		Efectivo,               ")
            loComandoSeleccionar.AppendLine(" 		Ticket,                 ")
            loComandoSeleccionar.AppendLine(" 		Cheque,                 ")
            loComandoSeleccionar.AppendLine(" 		Tarjeta,                ")
            loComandoSeleccionar.AppendLine(" 		Depósito,               ")
            loComandoSeleccionar.AppendLine(" 		Transferencia,          ")
            loComandoSeleccionar.AppendLine(" 		Total_Cobros,           ")
            loComandoSeleccionar.AppendLine(" 		UltimaSemana, ")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YYYY , DATEADD (ww , -46, " & lcParametro0Hasta & " )) As Menor_Año,   ")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YYYY , " & lcParametro0Hasta & " ) As Mayor_Año,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(ww , DATEADD (ww , -46, " & lcParametro0Hasta & " )) As Parametro_Semana_Inicio,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(ww , " & lcParametro0Hasta & " ) As Parametro_Semana_Fin")
            loComandoSeleccionar.AppendLine(" FROM #Final")
            loComandoSeleccionar.AppendLine(" ORDER BY Año,Semana Asc")
            loComandoSeleccionar.AppendLine(" ")



			'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

	'******************************************************************************************
	' Inicio Se Procesa manualmetne los datos
	'******************************************************************************************

		'Tabla con las listas desplegables
		Dim loTabla As New DataTable("curReportes")
		Dim loColumna As DataColumn 
		
		loColumna = New DataColumn("Año", getType(integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Semana", getType(integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Efectivo", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Ticket", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Cheque", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Tarjeta", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Depósito", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Transferencia", getType(decimal))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Total_Cobros", getType(decimal))
		loTabla.Columns.Add(loColumna)
		

	   IF laDatosReporte.Tables(0).Rows.Count > 0 Then
	
					Dim loNuevaFila As DataRow
					Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
						
					Dim PrimerAño As Integer = laDatosReporte.Tables(0).Rows(0).Item("Menor_Año")
					Dim SegundoAño As Integer = laDatosReporte.Tables(0).Rows(0).Item("Mayor_Año")
					Dim PrimeraSemana As Integer = laDatosReporte.Tables(0).Rows(0).Item("Semana")
					Dim DiaSegundoAño As Integer = 1
					
					
					
					Dim Semana_Inicio As Integer = laDatosReporte.Tables(0).Rows(0).Item("Parametro_Semana_Inicio")
					Dim Semana_Hasta As Integer =  laDatosReporte.Tables(0).Rows(0).Item("Parametro_Semana_Fin") 

					
					'******************************************************************************************
					' Se inicializa una tabla desde la primera hasta la ultima semana solicitada por parametros
					' Tomando en cuenta si las primeras semanas son de un año y las ultimas del año proximo, ó 
					' si todas las semanas son de un mismo año
					'******************************************************************************************
					
					
					'******************************************************************************************
					' En Caso que hubieran semanas de dos años
					'******************************************************************************************
					IF laDatosReporte.Tables(0).Rows(0).Item("Año") < laDatosReporte.Tables(0).Rows(lnTotalFilas-1).Item("Año") Then

						For lnNumeroFila As Integer = Semana_inicio To laDatosReporte.Tables(0).Rows(0).Item("UltimaSemana") 
				
							loNuevaFila = loTabla.NewRow()
							loTabla.Rows.Add(loNuevaFila)
							

								loNuevaFila.Item("Año")					= PrimerAño
								loNuevaFila.Item("Semana")				= lnNumeroFila
								loNuevaFila.Item("Efectivo")			= 0.0
								loNuevaFila.Item("Ticket")				= 0.0
								loNuevaFila.Item("Cheque")				= 0.0
								loNuevaFila.Item("Tarjeta")				= 0.0
								loNuevaFila.Item("Depósito")			= 0.0
								loNuevaFila.Item("Transferencia")		= 0.0
								loNuevaFila.Item("Total_Cobros")		= 0.0	
								
								
								loTabla.AcceptChanges()
								
								'PrimeraSemana = PrimeraSemana + 1

						Next
						

						For lnNumeroFila As Integer = 1 To Semana_Hasta 
															
							loNuevaFila = loTabla.NewRow()
							loTabla.Rows.Add(loNuevaFila)
							

								loNuevaFila.Item("Año")					= laDatosReporte.Tables(0).Rows(lnTotalFilas-1).Item("Año") 
								loNuevaFila.Item("Semana")				= lnNumeroFila
								loNuevaFila.Item("Efectivo")			= 0.0
								loNuevaFila.Item("Ticket")				= 0.0
								loNuevaFila.Item("Cheque")				= 0.0
								loNuevaFila.Item("Tarjeta")				= 0.0
								loNuevaFila.Item("Depósito")			= 0.0
								loNuevaFila.Item("Transferencia")		= 0.0
								loNuevaFila.Item("Total_Cobros")		= 0.0	
								
								loTabla.AcceptChanges()

						Next

					End If
			
					
					'******************************************************************************************
					' En Caso que no hubieran semanas de dos años
					'******************************************************************************************
					
					IF laDatosReporte.Tables(0).Rows(0).Item("Año") = laDatosReporte.Tables(0).Rows(lnTotalFilas-1).Item("Año") Then

						For lnNumeroFila As Integer = Semana_Inicio To Semana_Hasta
				
							loNuevaFila = loTabla.NewRow()
							loTabla.Rows.Add(loNuevaFila)
							

								loNuevaFila.Item("Año")					= PrimerAño
								loNuevaFila.Item("Semana")				= lnNumeroFila
								loNuevaFila.Item("Efectivo")			= 0.0
								loNuevaFila.Item("Ticket")				= 0.0
								loNuevaFila.Item("Cheque")				= 0.0
								loNuevaFila.Item("Tarjeta")				= 0.0
								loNuevaFila.Item("Depósito")			= 0.0
								loNuevaFila.Item("Transferencia")		= 0.0
								loNuevaFila.Item("Total_Cobros")		= 0.0	
								
								loTabla.AcceptChanges()

						Next

					End If
			
					'******************************************************************************************
					' Se evalua la tabla con la data del select y se rellena la tabla inicializada anteriormente
					'******************************************************************************************

					For Each loRenglonActual as DataRow in laDatosReporte.Tables(0).Rows

						Dim Renglon As DataRow  =loTabla.NewRow()
						Dim LcAxuAño As String = loRenglonActual.Item("Año").ToString
						Dim LcAxuSemana As String = loRenglonActual.Item("Semana").ToString	
							
							Renglon.Item("Año")				= LcAxuAño
							Renglon.Item("Semana")			= LcAxuSemana
							Renglon.Item("Efectivo")			= loRenglonActual.Item("Efectivo")
							Renglon.Item("Ticket")			= loRenglonActual("Ticket")
							Renglon.Item("Cheque")			= loRenglonActual("Cheque")
							Renglon.Item("Tarjeta")			= loRenglonActual("Tarjeta")		
							Renglon.Item("Depósito")			= loRenglonActual("Depósito")		
							Renglon.Item("Transferencia")	= loRenglonActual("Transferencia")	
							Renglon.Item("Total_Cobros")		= loRenglonActual("Total_Cobros")
							loTabla.Rows.Add(Renglon)	

					Next
		Else 
				Dim SemanaFin As Integer = Microsoft.VisualBasic.Format(DatePART(DateInterval.WeekOfYear, cusAplicacion.goReportes.paParametrosFinales(0)))
				Dim Año As Integer = Microsoft.VisualBasic.Format(DatePART(DateInterval.Year, cusAplicacion.goReportes.paParametrosFinales(0)))
				
				
				Dim loNuevaFila As DataRow
				Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
				
				If SemanaFin - 46 <= 0 Then

					For lnNumeroFila As Integer = Math.Abs(SemanaFin - 46) To 54
														
						loNuevaFila = loTabla.NewRow()
						loTabla.Rows.Add(loNuevaFila)

							loNuevaFila.Item("Año")					= (Año - 1)
							loNuevaFila.Item("Semana")				= lnNumeroFila
							loNuevaFila.Item("Efectivo")			= 0.0
							loNuevaFila.Item("Ticket")				= 0.0
							loNuevaFila.Item("Cheque")				= 0.0
							loNuevaFila.Item("Tarjeta")				= 0.0
							loNuevaFila.Item("Depósito")			= 0.0
							loNuevaFila.Item("Transferencia")		= 0.0
							loNuevaFila.Item("Total_Cobros")		= 0.0	
							
							loTabla.AcceptChanges()

					Next
					
					For lnNumeroFila As Integer = 1 To SemanaFin
			
						loNuevaFila = loTabla.NewRow()
						loTabla.Rows.Add(loNuevaFila)
						

							loNuevaFila.Item("Año")					= Año
							loNuevaFila.Item("Semana")				= lnNumeroFila
							loNuevaFila.Item("Efectivo")			= 0.0
							loNuevaFila.Item("Ticket")				= 0.0
							loNuevaFila.Item("Cheque")				= 0.0
							loNuevaFila.Item("Tarjeta")				= 0.0
							loNuevaFila.Item("Depósito")			= 0.0
							loNuevaFila.Item("Transferencia")		= 0.0
							loNuevaFila.Item("Total_Cobros")		= 0.0	
							
							loTabla.AcceptChanges()
					Next

				Else
				
					For lnNumeroFila As Integer = Math.Abs(SemanaFin - 46) To SemanaFin
		
					loNuevaFila = loTabla.NewRow()
					loTabla.Rows.Add(loNuevaFila)
					

						loNuevaFila.Item("Año")					= Año
						loNuevaFila.Item("Semana")				= lnNumeroFila
						loNuevaFila.Item("Efectivo")			= 0.0
						loNuevaFila.Item("Ticket")				= 0.0
						loNuevaFila.Item("Cheque")				= 0.0
						loNuevaFila.Item("Tarjeta")				= 0.0
						loNuevaFila.Item("Depósito")			= 0.0
						loNuevaFila.Item("Transferencia")		= 0.0
						loNuevaFila.Item("Total_Cobros")		= 0.0	
						
						loTabla.AcceptChanges()

					Next
				
				End If
				
		End If


	'******************************************************************************************
	' Fin Se Procesa manualmetne los datos
	'******************************************************************************************

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

			dim loTablaRellenada As New DataSet()
			'me.mEscribirConsulta(loTablaRellenada.ToString)
			loTablaRellenada.Tables.Add(loTabla)
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_Semanalmente", loTablaRellenada)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCobros_Semanalmente.ReportSource = loObjetoReporte

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
' CMS: 05/05/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 03/02/11: Ajuste del Select, no mostraba información.
'-------------------------------------------------------------------------------------------'
