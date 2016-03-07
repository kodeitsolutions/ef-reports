'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rABC_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rABC_Clientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))


            Dim loComandoSeleccionar As New StringBuilder()

			
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		ABC AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'ABC' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" INTO #temp ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Abc ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY ABC ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Status AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'ESTMAE' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Status ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Status ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Sexo AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'SEXO' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Sexo ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Sexo ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Crm_Sec AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'SECECO' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Crm_Sec ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Crm_Sec ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Tip_Con AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'TIPCON' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Tip_Con ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Tip_Con ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Pol_Com AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'POLCOM' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Pol_Com ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Pol_Com ")
            loComandoSeleccionar.AppendLine(" UNION  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		COUNT(Cod_Cli) AS Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Tip_Cli AS Valor, ")
            loComandoSeleccionar.AppendLine(" 		'NORSUCMAT' AS Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM Clientes AS Clientes_Tip_Cli ")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 				Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Tip_Cli ")
            loComandoSeleccionar.AppendLine(" ORDER BY 3, 2 ")
            loComandoSeleccionar.AppendLine("  ")
            loComandoSeleccionar.AppendLine("  ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 	Can_Cli, ")
            loComandoSeleccionar.AppendLine(" 	Valor, ")
            loComandoSeleccionar.AppendLine(" 	Cod_Com ")
            loComandoSeleccionar.AppendLine(" FROM #temp ")
			loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Com ")
            loComandoSeleccionar.AppendLine("  ")
            loComandoSeleccionar.AppendLine(" SELECT * FROM Factory_Global.dbo.Combos  ")
            loComandoSeleccionar.AppendLine(" WHERE Cod_Com IN ('ABC', 'ESTMAE', 'Sexo', 'SECECO', 'TIPCON', 'POLCOM', 'NORSUCMAT') ")
            loComandoSeleccionar.AppendLine(" ORDER BY Cod_Com ")

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
    '******************************************************************************************
	' Inicio Se Procesa manualmetne los datos
	'******************************************************************************************

		'Tabla con las listas desplegables
		Dim loTabla As New DataTable("curReportes")
		Dim loColumna As DataColumn 
		
		loColumna = New DataColumn("Cod_Com", getType(String))
		loColumna.MaxLength = 50
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Valor", getType(String))
		loColumna.MaxLength = 300
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("can_cli", GetType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("porcentaje", GetType(Integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("texto", getType(String))
		loColumna.MaxLength = 300
		loTabla.Columns.Add(loColumna)
		

		Dim loNuevaFila As DataRow
		Dim lnTotalFilas AS Integer = laDatosReporte.Tables(1).Rows.Count
		
		For lnNumeroFila As Integer = 0 To lnTotalFilas - 1  
			
			'Extrayendo los valores de la lista
			Dim lcXml As String = "<elementos></elementos>"
            Dim lcItem As String
            Dim lnItemNumero As Integer = 0
            Dim loItems As New System.Xml.XmlDocument()

            lcItem = ""
     
              lcXml = laDatosReporte.Tables(1).Rows(lnNumeroFila).Item("contenido")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loItems.LoadXml(lcXml)

                'Extraccion de los elementos de una lista desplegable
                For Each loItem As System.Xml.XmlNode In loItems.SelectNodes("elementos/elemento")

                    If lnNumeroFila <= laDatosReporte.Tables(1).Rows.Count - 1 Then
                        
						Dim loFila As DataRow = laDatosReporte.Tables(1).Rows(lnNumeroFila)
						loNuevaFila = loTabla.NewRow()
						loTabla.Rows.Add(loNuevaFila)

						lnItemNumero = lnItemNumero + 1
						
                        loNuevaFila.Item("cod_com")			= Trim(loFila("cod_com"))
						loNuevaFila.Item("valor")			= Trim(loItem.Attributes("valor").Value).ToString
						loNuevaFila.Item("Texto")			= loItem.InnerText.ToString
						loNuevaFila.Item("can_cli")			=  0.0
						loNuevaFila.Item("porcentaje")		=  0.0
						
						loTabla.AcceptChanges()	                        
                    End If
                Next loItem
                
                lnItemNumero = 0

		Next lnNumeroFila
		

		'Tabla con los datos finales
		Dim loTablaFinal As New DataTable("curReportes")
		Dim loColumnaFinal As DataColumn 
		
		loColumnaFinal = New DataColumn("Cod_Com", getType(String))
		loColumnaFinal.MaxLength = 50
		loTablaFinal.Columns.Add(loColumnaFinal)
		
		loColumnaFinal = New DataColumn("Valor", getType(String))
		loColumnaFinal.MaxLength = 300
		loTablaFinal.Columns.Add(loColumnaFinal)
		
		loColumnaFinal = New DataColumn("can_cli", GetType(Integer))
		loTablaFinal.Columns.Add(loColumnaFinal)
		
		loColumnaFinal = New DataColumn("porcentaje", GetType(Integer))
		loTablaFinal.Columns.Add(loColumnaFinal)
		

		Dim lnTotalFilas2 AS Integer = laDatosReporte.Tables(0).Rows.Count
		
		
		For i As Integer = 0 To lnTotalFilas2 - 1
		
			For j As Integer = 0 To loTabla.Rows.Count - 1

				IF Trim(laDatosReporte.Tables(0).Rows(i).Item(2).ToString) = Trim(loTabla.Rows(j).Item(0).ToString) _
					AND Trim(laDatosReporte.Tables(0).Rows(i).Item(1).ToString) = Trim(loTabla.Rows(j).Item(1).ToString) Then
					
					loTabla.Rows(j).Item(0)		= laDatosReporte.Tables(0).Rows(i).Item(2).ToString
					loTabla.Rows(j).Item(1)		= laDatosReporte.Tables(0).Rows(i).Item(1).ToString
					loTabla.Rows(j).Item(2)		= laDatosReporte.Tables(0).Rows(i).Item(0).ToString
					loTabla.Rows(j).Item(3)		= 0.0
					
					loTabla.AcceptChanges()
					
				End If

			Next j

		Next i	
		
		Dim loDatosReporteFinal As New DataSet("curReportes") 
		loDatosReporteFinal.Tables.Add(loTabla)


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

            'loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rABC_Clientes", laDatosReporte)
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rABC_Clientes", loDatosReporteFinal)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrABC_Clientes.ReportSource = loObjetoReporte

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
' CMS: 14/04/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'