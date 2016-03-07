'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grIncidencias_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class grIncidencias_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            'If lcParametro3Desde = "" Then
            'lcParametro3Desde = ""
            'lcParametro3Hasta = "zzzzzzzz"
            'Else
            'lcParametro3Hasta = lcParametro3Desde
            'End If

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            'If lcParametro4Desde = "" Then
            'lcParametro4Desde = ""
            'lcParametro4Hasta = "zzzzzzzz"
            'Else
            'lcParametro4Hasta = lcParametro4Desde
            'End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" DECLARE @Numero decimal")


            'loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Cobros.Fec_Ini, 103)	AS	Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Cobros.Fec_Ini, 112)	AS	Fecha2,	")



            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Registro)            AS  Año, ")
            loComandoSeleccionar.AppendLine(" 	        CONVERT(NCHAR(10), Registro, 112)	AS	Fecha, ")
            loComandoSeleccionar.AppendLine(" 	        CONVERT(NCHAR(10), Registro, 103)	AS	Fecha2, ")
            loComandoSeleccionar.AppendLine("         	Can_Err                             AS  Cantidad ")
            loComandoSeleccionar.AppendLine(" INTO      #Temp ")
            loComandoSeleccionar.AppendLine(" FROM      Factory_Global.dbo.Errores ")
            loComandoSeleccionar.AppendLine(" WHERE     Registro                Between DATEADD (MONTH , -1, "& lcParametro0Hasta &" )" ) 
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Usu             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Status              IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Sistema             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Modulo              Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine(" SELECT    Año,")
            loComandoSeleccionar.AppendLine(" 	        SUBSTRING(CAST(Fecha As VARCHAR),5,4)   AS Fecha,")
            loComandoSeleccionar.AppendLine(" 	        Cantidad                                AS Cantidad ")
            loComandoSeleccionar.AppendLine(" INTO      #Temp2")
            loComandoSeleccionar.AppendLine(" FROM      #Temp")

            loComandoSeleccionar.AppendLine(" SELECT    Año,")
            loComandoSeleccionar.AppendLine(" 	        Fecha,")
            loComandoSeleccionar.AppendLine(" 	        Fecha2,")
            loComandoSeleccionar.AppendLine(" 	        SUM(Cantidad) AS Cantidad")
            loComandoSeleccionar.AppendLine(" FROM      #Temp")
            loComandoSeleccionar.AppendLine(" GROUP BY  Año,Fecha, Fecha2")
            loComandoSeleccionar.AppendLine(" ORDER BY  Año,Fecha")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grIncidencias_Fechas", laDatosReporte)




	'******************************************************************************************
	' Inicio Se Procesa manualmetne los datos
	'******************************************************************************************

		'Tabla con las listas desplegables
		Dim loTabla As New DataTable("curReportes")
		Dim loColumna As DataColumn 
		
		loColumna = New DataColumn("Año", getType(integer))
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Fecha", getType(string))
		'loColumna.MaxLength = 8
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Fecha2", getType(string))
		'loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("Cantidad", getType(decimal))
		loTabla.Columns.Add(loColumna)


	   IF laDatosReporte.Tables(0).Rows.Count > 0 Then
	
			Dim loNuevaFila As DataRow
			'Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
				
			Dim Dia_Inicio As Date	= Microsoft.VisualBasic.Format(CDate(cusAplicacion.goReportes.paParametrosFinales(0)).AddMonths(-1))
			Dim Dia_Fin As Date		= Microsoft.VisualBasic.Format(CDate(cusAplicacion.goReportes.paParametrosFinales(0)))
			 
			
			While  Dia_Inicio <= Dia_Fin 
			
					loNuevaFila = loTabla.NewRow()
					loTabla.Rows.Add(loNuevaFila)
					

						loNuevaFila.Item("Año")					= Dia_Inicio.Year
						loNuevaFila.Item("Fecha")				= Microsoft.VisualBasic.Format(Dia_Inicio, "yyyyMMdd")
						loNuevaFila.Item("Fecha2")				= Microsoft.VisualBasic.Format(Dia_Inicio,"dd/MM/yyyy" )
						loNuevaFila.Item("Cantidad")			= 0.0
						
						
						loTabla.AcceptChanges()

						Dia_Inicio = Dia_Inicio.AddDays(1) 
						
			End While
			
			
			For Each loRenglonActual as DataRow in laDatosReporte.Tables(0).Rows

				Dim Renglon As DataRow
				Dim LcAxuFecha As String = loRenglonActual.Item("Fecha2").ToString
						
					LcAxuFecha = Trim(LcAxuFecha)
					Renglon = loTabla.Select("Trim(Fecha2) = '" & Microsoft.VisualBasic.Format("dd/MM/yyyy", LcAxuFecha) & "'") (0)
					
					Renglon.Item("Año")					= loRenglonActual("Año")
					Renglon.Item("Fecha")				= loRenglonActual("Fecha")
					Renglon.Item("Fecha2")				= loRenglonActual("Fecha2")
					Renglon.Item("Cantidad")			= loRenglonActual("Cantidad")

			Next
	
		Else
		
		
			Dim loNuevaFila As DataRow
				
			Dim Dia_Inicio As Date	= Microsoft.VisualBasic.Format(CDate(cusAplicacion.goReportes.paParametrosFinales(0)).AddMonths(-1))
			Dim Dia_Fin As Date		= Microsoft.VisualBasic.Format(CDate(cusAplicacion.goReportes.paParametrosFinales(0)))
			
			While  Dia_Inicio <= Dia_Fin 
			
					loNuevaFila = loTabla.NewRow()
					loTabla.Rows.Add(loNuevaFila)
					

						loNuevaFila.Item("Año")					= Dia_Inicio.Year
						loNuevaFila.Item("Fecha")				= Microsoft.VisualBasic.Format(Dia_Inicio, "yyyyMMdd")
						loNuevaFila.Item("Fecha2")				= Microsoft.VisualBasic.Format(Dia_Inicio,"dd/MM/yyyy" )
						loNuevaFila.Item("Cantidad")			= 0.0
						
						
						loTabla.AcceptChanges()

						Dia_Inicio = Dia_Inicio.AddDays(1) 
						
			End While
		
		
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


			Dim loTablaRellenada As New DataSet()
			loTablaRellenada.Tables.Add(loTabla)
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grIncidencias_Fechas", loTablaRellenada)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrIncidencias_Fechas.ReportSource = loObjetoReporte

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
' JJD: 07/11/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 21/07/10: Se limito el numero de fechas a mostrar en el reporte y se relleno con cero
'				 las fechas donde no huebieran incidencias
'-------------------------------------------------------------------------------------------'