'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrestamos_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rPrestamos_Numeros 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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

			Dim loConsulta As New StringBuilder()




            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Prestamos.Documento,     ")
            loConsulta.AppendLine("        Prestamos.Status,       ")
            loConsulta.AppendLine("        Prestamos.cod_tra,       ")
            loConsulta.AppendLine("        Trabajadores.nom_tra,    ")
            loConsulta.AppendLine("        Prestamos.cod_tip,       ")
            loConsulta.AppendLine("        Prestamos.fec_asi,       ")
            loConsulta.AppendLine("        Prestamos.mon_net,       ")
            loConsulta.AppendLine("        Prestamos.mon_sal,       ")
            loConsulta.AppendLine("        Prestamos.mon_pag,       ")
            loConsulta.AppendLine("        Prestamos.comentario     ")
            loConsulta.AppendLine("FROM    Prestamos")
            loConsulta.AppendLine("JOIN	Trabajadores")
            loConsulta.AppendLine("	ON	Trabajadores.Cod_Tra = Prestamos.Cod_Tra")
            loConsulta.AppendLine("    AND Prestamos.Documento  BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Prestamos.Fec_Asi    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND Prestamos.Cod_Tra    BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Prestamos.Status      IN (" & lcParametro3Desde & ")")
            loConsulta.AppendLine("    AND Prestamos.Cod_Tip    BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("    AND " & lcParametro4Hasta)
            loConsulta.AppendLine("    AND Prestamos.Cod_Rev    BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("    AND " & lcParametro5Hasta)
            loConsulta.AppendLine("    AND Trabajadores.Cod_Con   BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")



            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rPrestamos_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPrestamos_Numeros.ReportSource =	 loObjetoReporte	


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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 03/04/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
