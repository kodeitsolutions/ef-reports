'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rConceptosNomina_ConSusGrupos"
'-------------------------------------------------------------------------------------------'
Partial Class rConceptosNomina_ConSusGrupos
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      

	Try	
	
	    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
	    Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
	    Dim lcParametro4Desde As String = cusAplicacion.goReportes.paParametrosIniciales(4)
	    Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

        Dim llIncluirConceptosSinAsignar As Boolean = (lcParametro4Desde.ToUpper() = "SI")

		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loConsulta As New StringBuilder()
	
	 
			loConsulta.AppendLine("SELECT      Conceptos_Nomina.cod_con                 AS Cod_Con,")
			loConsulta.AppendLine("            Conceptos_Nomina.nom_con                 AS Nom_Con,")
			loConsulta.AppendLine("            Conceptos_Nomina.tipo                    AS Tipo,")
			loConsulta.AppendLine("            Conceptos_Nomina.clase                   AS Clase,")
			loConsulta.AppendLine("            (CASE WHEN Conceptos_Nomina.Status='A'")
			loConsulta.AppendLine("                THEN 'Activo' ELSE 'Inactivo'")
			loConsulta.AppendLine("            END)                                     AS Estatus_Concepto,")
			loConsulta.AppendLine("            Grupos_Conceptos.Cod_Gru                 AS Cod_Gru,")
			loConsulta.AppendLine("            Grupos_Conceptos.nom_gru                 AS Nom_Gru,")
			loConsulta.AppendLine("            (CASE COALESCE(Grupos_Conceptos.Status, '')")
			loConsulta.AppendLine("                WHEN 'A' THEN 'Activo'")
			loConsulta.AppendLine("                WHEN '' THEN NULL")
			loConsulta.AppendLine("                ELSE 'Inactivo'")
			loConsulta.AppendLine("            END)                                     AS Estatus_Grupo ")
			loConsulta.AppendLine("FROM        Conceptos_Nomina")
			loConsulta.AppendLine("    LEFT JOIN Renglones_Grupos_Conceptos")
			loConsulta.AppendLine("        ON  Renglones_Grupos_Conceptos.cod_con = Conceptos_Nomina.cod_con")
			loConsulta.AppendLine("    LEFT JOIN Grupos_Conceptos")
			loConsulta.AppendLine("        ON  Grupos_Conceptos.Cod_Gru = Renglones_Grupos_Conceptos.cod_gru")
			loConsulta.AppendLine("        AND Grupos_Conceptos.Cod_Gru BETWEEN " & lcParametro5Desde )
			loConsulta.AppendLine("            AND " & lcParametro5Hasta )
			loConsulta.AppendLine("WHERE       Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde )
			loConsulta.AppendLine("            AND " & lcParametro0Hasta )
			loConsulta.AppendLine("        AND Conceptos_Nomina.Tipo IN (" & lcParametro1Desde & ")" )
			loConsulta.AppendLine("        AND Conceptos_Nomina.Clase IN (" & lcParametro2Desde & ")" )
			loConsulta.AppendLine("        AND Conceptos_Nomina.Status IN (" & lcParametro3Desde & ")" )
            If Not llIncluirConceptosSinAsignar Then 
			    loConsulta.AppendLine("        AND Grupos_Conceptos.Cod_Gru IS NOT NULL" )
            End If
			loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento )
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")


            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos()
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConceptosNomina_ConSusGrupos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrConceptosNomina_ConSusGrupos.ReportSource = loObjetoReporte
			
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
' RJG: 28/11/13: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
