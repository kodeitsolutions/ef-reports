'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rCampos_Nomina"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rCampos_Nomina
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	    Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		    Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCodCam_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodCam_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Campos_Nomina.Cod_Cam,")
            loConsulta.AppendLine("		    Campos_Nomina.Nom_Cam,")
            loConsulta.AppendLine("		    Campos_Nomina.Uso,")
            loConsulta.AppendLine("		    Renglones_Campos_Nomina.Val_Num,")
            loConsulta.AppendLine("		    Trabajadores.Cod_Tra,")
            loConsulta.AppendLine("		    Trabajadores.Nom_Tra")
            loConsulta.AppendLine("FROM Campos_Nomina")
            loConsulta.AppendLine("	JOIN Renglones_Campos_Nomina ON Campos_Nomina.Cod_Cam = Renglones_Campos_Nomina.Cod_Cam")
            loConsulta.AppendLine("	JOIN Trabajadores ON Renglones_Campos_Nomina.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("		AND Trabajadores.Status = 'A'")
            loConsulta.AppendLine("WHERE Campos_Nomina.Cod_Cam BETWEEN @lcCodCam_Desde AND @lcCodCam_Hasta")
            loConsulta.AppendLine("     AND Campos_Nomina.Tipo = 'N'")
            loConsulta.AppendLine("     AND Campos_Nomina.Status = 'A'")
            loConsulta.AppendLine("	 AND Renglones_Campos_Nomina.Val_Num > 0")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

           
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rCampos_Nomina", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvTRF_rCampos_Nomina.ReportSource = loObjetoReporte
			  
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
