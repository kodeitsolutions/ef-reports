'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIncidencias_Accesos"
'-------------------------------------------------------------------------------------------'
Partial Class rIncidencias_Accesos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

            Dim loComandoSeleccionar As New StringBuilder()

            Dim ldInicio	As DateTime	= cusAplicacion.goReportes.paParametrosIniciales(1)
			Dim ldFin		As DateTime	= cusAplicacion.goReportes.paParametrosFinales(1)
			
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0)	)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0)	)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(ldInicio, goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(ldFin, goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)	)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2)	)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3)	)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3)	)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine(" SELECT	Incidencias.Documento,")
            loComandoSeleccionar.AppendLine("			Incidencias.Registro AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			''		AS Fecha,")
            loComandoSeleccionar.AppendLine("			Incidencias.Cod_Tra,")
            loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra,")
            loComandoSeleccionar.AppendLine("			ISNULL(Trabajadores.Cod_Dep, '*') AS Cod_Dep,")
            loComandoSeleccionar.AppendLine("			ISNULL(Departamentos_Nomina.Nom_Dep, '*No Asignado*') AS Nom_Dep,")
            loComandoSeleccionar.AppendLine("			Incidencias.Tipo, ")
            loComandoSeleccionar.AppendLine("			Incidencias.Notas, " )
            loComandoSeleccionar.AppendLine("			Incidencias.Comentario " )
            loComandoSeleccionar.AppendLine(" FROM		Incidencias ")
            loComandoSeleccionar.AppendLine("	        LEFT JOIN Trabajadores  ON Trabajadores.Cod_Tra  = Incidencias.Cod_Tra")
            loComandoSeleccionar.AppendLine("	        LEFT JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            loComandoSeleccionar.AppendLine(" WHERE     Trabajadores.Status = 'A'")
            loComandoSeleccionar.AppendLine("		    AND	Incidencias.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		    AND	Incidencias.Registro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		    AND	Incidencias.Cod_Tra BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND	" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		    AND	Incidencias.Cod_Dep BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND	" & lcParametro3Hasta)
            'loComandoSeleccionar.AppendLine("ORDER BY	Incidencias.Cod_Dep, Incidencias.Cod_Tra, Incidencias.Fec_Ini")
            loComandoSeleccionar.AppendLine("ORDER BY   Incidencias.Cod_Dep, " & lcOrdenamiento)
            
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
   
            Dim loServicios As New cusDatos.goDatos
			Dim laDatosReporte As DataSet 
			
        Try
										    
            laDatosReporte = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			
        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")
			Return 
			
        End Try




		For Each loRegistro AS DataRow in laDatosReporte.Tables(0).Rows
			
			loRegistro("Fecha") = goServicios.mObtenerFormatoCadena(loRegistro("Fec_Ini"), _
									goServicios.enuOpcionesRedondeo.KN_FechaSinDecimales)
									
			
		Next loRegistro

		laDatosReporte.AcceptChanges()





    
	''-------------------------------------------------------------------------------------------'
	''  Crea el contenedor para las inasistencias.												'
	''-------------------------------------------------------------------------------------------'
		
	'	Dim loTablaReporte As New DataTable("curReportes")
		
	'	loTablaReporte.Columns.Add("Fecha",		GetType(String))
	'	loTablaReporte.Columns.Add("dia",		GetType(String))
	'	loTablaReporte.Columns.Add("cod_tra", 	GetType(String))
	'	loTablaReporte.Columns.Add("nom_tra", 	GetType(String))
	'	loTablaReporte.Columns.Add("cod_dep", 	GetType(String))
	'	loTablaReporte.Columns.Add("nom_dep", 	GetType(String))
	'	loTablaReporte.Columns.Add("cod_tar", 	GetType(String))
		
		
	''-------------------------------------------------------------------------------------------'
	''  Recorre la lista de trabajadores que no vinieron en el rango de fechas.					'
	''-------------------------------------------------------------------------------------------'
	'	Dim lnTotalDias As Integer = Math.Ceiling((ldFin - ldInicio).TotalDays)
	'	Dim ldDiaInicial As DateTime = New DateTime(ldInicio.Year, ldInicio.Month, ldInicio.Day)
		
	'	For Each loInasistente As DataRow In laDatosReporte.Tables("Inasistentes").Rows
		
	'		Dim lcTrabajador As String = goServicios.mObtenerCampoFormatoSQL(loInasistente("Cod_Tra"))
			
	'	'-------------------------------------------------------------------------------------------'
	'	'  Para cada inasistente, verifica cada día del rango de fechas y verifica si ese día vino.	'
	'	'-------------------------------------------------------------------------------------------'
	'		For lnDiaActual As Integer = 0 To lnTotalDias - 1	
				
	'			Dim ldDiaActual As DateTime = ldDiaInicial.AddDays(lnDiaActual)
				
	'			Dim lcSeleccion As String = "Cod_Tra=" & lcTrabajador  & " AND Registro >= '" & _
	'									 	Microsoft.VisualBasic.Format(ldDiaActual, "MM-dd-yyyy HH:mm:ss") & "' AND Registro <= '" & _
	'									 	Microsoft.VisualBasic.Format(ldDiaActual.AddDays(1).AddSeconds(-1), "MM-dd-yyyy HH:mm:ss") & "'"
				
	'			If laDatosReporte.Tables("Asistencias").Select(lcSeleccion).Length <= 0 Then 
	'				Dim loNuevoRegistro As DataRow = loTablaReporte.NewRow()
					
	'				loNuevoRegistro("Fecha")	= goServicios.mObtenerFormatoCadena(ldDiaActual, goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
	'				loNuevoRegistro("cod_tra")	= loInasistente("cod_tra")
	'				loNuevoRegistro("nom_tra")	= loInasistente("nom_tra")
	'				loNuevoRegistro("cod_dep")	= loInasistente("cod_dep")
	'				loNuevoRegistro("nom_dep")	= loInasistente("nom_dep")
	'				loNuevoRegistro("cod_tar")	= loInasistente("cod_tar")
					
	'				Select Case ldDiaActual.DayOfWeek 
	'					Case DayOfWeek.Monday
	'						loNuevoRegistro("dia") = "LUN"
	'					Case DayOfWeek.Tuesday
	'						loNuevoRegistro("dia") = "MAR"
	'					Case DayOfWeek.Wednesday
	'						loNuevoRegistro("dia") = "MIE"
	'					case DayOfWeek.Thursday 
	'						loNuevoRegistro("dia") = "JUE"
	'					case DayOfWeek.Friday
	'						loNuevoRegistro("dia") = "VIE"
	'					Case DayOfWeek.Saturday
	'						loNuevoRegistro("dia") = "SAB"
	'					Case DayOfWeek.Sunday
	'						loNuevoRegistro("dia") = "DOM"
	'				End Select
					
	'				loTablaReporte.Rows.Add(loNuevoRegistro)
					
	'			End If
				
	'		Next lnDiaActual
			
	'	Next loInasistente
			
		Try
			'Dim loDataSetReporte As New DataSet() 
			'loDataSetReporte.Tables.Add(loTablaReporte)
			
		    loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIncidencias_Accesos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrIncidencias_Accesos.ReportSource = loObjetoReporte


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
' MVP: 14/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 24/07/08: Se suprimión el filtro de la puertas puesto que inicialmente no se trabaja_
'                 con este dato necesario al ingresar o salir de la empresa.
'-------------------------------------------------------------------------------------------'
' CMS:  19/02/10: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'