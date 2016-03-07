'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTiempos_Documentos"
'-------------------------------------------------------------------------------------------'
Partial Class rTiempos_Documentos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            Dim loComandoValidar As New StringBuilder()

			loComandoValidar.AppendLine("SELECT	[Name]")
			loComandoValidar.AppendLine("FROM		Sys.Objects")
			loComandoValidar.AppendLine("WHERE		[Type]='U'")           
			loComandoValidar.AppendLine("			AND	[Name]='" & lcParametro1Desde.Replace("'","") & "'")
			
			Dim loServicios1 As New cusDatos.goDatos
			Dim laDatosReporteValido As DataSet = loServicios1.mObtenerTodosSinEsquema(loComandoValidar.ToString, "curReportes")
			
			 '-------------------------------------------------------------------------------------------------------
			 ' Verificando si el select (Parametro 0) trae registros
			 '-------------------------------------------------------------------------------------------------------
			
			 If (laDatosReporteValido.Tables(0).Rows.Count <= 0) Then
			 	Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
			 				  "No se encontraron registros para los parámetros especificados,  por  favor  ingrese  un  nombre  válido  para  la  tabla.", _
			 				   vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
			 				   "350px", _
			 				   "200px")
			 	Return			   
			 		
			 End If								   			   

			loComandoSeleccionar.AppendLine("SELECT			" & lcParametro1Desde.Replace("'","") & ".Documento,")
			loComandoSeleccionar.AppendLine(" 				" & lcParametro1Desde.Replace("'","") & ".Fec_Ini,")
			loComandoSeleccionar.AppendLine(" 				" & lcParametro1Desde.Replace("'","") & ".Status,")            
			loComandoSeleccionar.AppendLine(" 				" & lcParametro1Desde.Replace("'","") & ".Reg_Ini,")
			loComandoSeleccionar.AppendLine(" 				" & lcParametro1Desde.Replace("'","") & ".Reg_Fin,")
			loComandoSeleccionar.AppendLine(" 				'     '						AS Tiempo,")
			loComandoSeleccionar.AppendLine(" 				CAST(0 AS INTEGER)			AS Minutos,")
			loComandoSeleccionar.AppendLine(" 				CAST(0 AS INTEGER)			AS Horas,")
			'loComandoSeleccionar.AppendLine(" 				CAST(0 AS INTEGER)			AS Miliseg,")
			loComandoSeleccionar.AppendLine(" 				CAST(0 AS INTEGER)			AS Segundos")
			loComandoSeleccionar.AppendLine("FROM			" & lcParametro1Desde.Replace("'","") & " ")
			loComandoSeleccionar.AppendLine("WHERE			" & lcParametro1Desde.Replace("'","") & ".Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("               AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("               AND " & lcParametro1Desde.Replace("'","")& ".Status IN (" & lcParametro2Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			
			Dim loServicios As New cusDatos.goDatos
			
					'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

					Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					
					For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows 
						 
						dim ldFechaInicio As Date = CDate(loRenglon("Reg_Ini"))
						dim ldFechaFin As Date = CDate(loRenglon("Reg_Fin"))
						
						Dim loDiferencia As TimeSpan = ldFechaFin - ldFechaInicio
						
						loRenglon("Tiempo") = Microsoft.VisualBasic.Format(loDiferencia.Minutes, "00")& ":" & Microsoft.VisualBasic.Format(loDiferencia.TotalSeconds, "00")
						loRenglon("Minutos") = Microsoft.VisualBasic.Format(loDiferencia.Minutes, "00")
						loRenglon("Horas") = Microsoft.VisualBasic.Format(loDiferencia.Hours, "00")
						'loRenglon("Miliseg") = Microsoft.VisualBasic.Format(loDiferencia.TotalMilliseconds, "00")
						loRenglon("Segundos") = Microsoft.VisualBasic.Format(loDiferencia.TotalSeconds, "00")			
					
					Next loRenglon
			

					loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTiempos_Documentos", laDatosReporte)
					
					Me.mTraducirReporte(loObjetoReporte)
		            
					Me.mFormatearCamposReporte(loObjetoReporte)

					Me.crvrTiempos_Documentos.ReportSource = loObjetoReporte

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
' MAT: 05/11/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 09/11/10: Programación del Tiempo Promedio
'-------------------------------------------------------------------------------------------'
' MAT: 30/04/11: Modificación para que tome el nombre de la tabla desde una lista desplegable
'-------------------------------------------------------------------------------------------'
