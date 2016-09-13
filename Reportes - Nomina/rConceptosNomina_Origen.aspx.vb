'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rConceptosNomina_Origen"
'-------------------------------------------------------------------------------------------'
Partial Class rConceptosNomina_Origen
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	    Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
	        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
	        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

		    Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		    Dim loConsulta As New StringBuilder()

		    loConsulta.AppendLine("SELECT   Cod_Con,")
		    loConsulta.AppendLine("         Nom_Con,")
		    loConsulta.AppendLine("         Status,")
		    loConsulta.AppendLine("         (CASE WHEN Status = 'A'")
		    loConsulta.AppendLine("             THEN 'Activo' ELSE 'Inactivo'")
		    loConsulta.AppendLine("          END) AS Estatus_Conceptos,")
		    loConsulta.AppendLine("         Tipo,")
		    loConsulta.AppendLine("         Tip_Ori,")
		    loConsulta.AppendLine("         Doc_Ori,")
		    loConsulta.AppendLine("         Ren_Ori")
		    loConsulta.AppendLine("FROM     Conceptos_Nomina")
		    loConsulta.AppendLine("WHERE    Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde )
		    loConsulta.AppendLine("         AND " & lcParametro0Hasta )
		    loConsulta.AppendLine("     AND Conceptos_Nomina.Status IN (" & lcParametro1Desde & ")")
		    loConsulta.AppendLine("     AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
		    loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
		    loConsulta.AppendLine("")
		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            'Los campos con código LIF (Tip_Ori, Doc_Ori, Ren_Ori...) vienen de base de datos codificados 
            'con goServicios.mCodificarQuotedPrintable()
            For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows
                
                loRenglon("Tip_Ori") = goServicios.mDecodificarQuotedPrintable(CStr(loRenglon("Tip_Ori")).Trim())
                loRenglon("Doc_Ori") = goServicios.mDecodificarQuotedPrintable(CStr(loRenglon("Doc_Ori")).Trim())
                loRenglon("Ren_Ori") = goServicios.mDecodificarQuotedPrintable(CStr(loRenglon("Ren_Ori")).Trim())

            Next loRenglon


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConceptosNomina_Origen", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrConceptosNomina_Origen.ReportSource = loObjetoReporte
			  
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
' RJG: 13/08/14: Codigo inicial.                                                           '
'-------------------------------------------------------------------------------------------'