'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rConceptos_Nomina"
'-------------------------------------------------------------------------------------------'
Partial Class rConceptos_Nomina
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	    Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
	        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

		    Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		    Dim loConsulta As New StringBuilder()

		    loConsulta.AppendLine("SELECT   Cod_Con,")
		    loConsulta.AppendLine("         Nom_Con,")
		    loConsulta.AppendLine("         Status,")
		    loConsulta.AppendLine("         (CASE WHEN Status = 'A'")
		    loConsulta.AppendLine("             THEN 'Activo' ELSE 'Inactivo'")
		    loConsulta.AppendLine("          END) AS Estatus_Conceptos")
		    loConsulta.AppendLine("FROM     Conceptos_Nomina")
		    loConsulta.AppendLine("WHERE    Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde )
		    loConsulta.AppendLine("         AND " & lcParametro0Hasta )
		    loConsulta.AppendLine("     AND Conceptos_Nomina.Status IN (" & lcParametro1Desde & ")")
		    loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
		    loConsulta.AppendLine("")
		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

           
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConceptos_Nomina", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrConceptos_Nomina.ReportSource = loObjetoReporte
			  
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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' MJP: 09/07/08 : Codigo inicial.                                                           '
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08 : Creación objeto que cierra el archivo de reporte.                         '
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08 : Agregacion filtro Status.                                                 '
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.                 '
'-------------------------------------------------------------------------------------------'
' RJG: 29/11/13: Se actualizó para apuntar a la nueva tabla "Conceptos_Nomina". Ajustes de  '
'                interfaz.                                                                  '
'-------------------------------------------------------------------------------------------'
