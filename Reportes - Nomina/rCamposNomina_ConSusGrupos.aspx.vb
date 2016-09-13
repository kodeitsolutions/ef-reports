'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCamposNomina_ConSusGrupos"
'-------------------------------------------------------------------------------------------'
Partial Class rCamposNomina_ConSusGrupos
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      

	Try	
	
	    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
	    Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
	    Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosIniciales(3)
	    Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

        Dim llIncluirCamposSinAsignar As Boolean = (lcParametro3Desde.ToUpper() = "SI")

		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loConsulta As New StringBuilder()
	
	 
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Campos_Nomina.Cod_Cam                   AS Cod_Cam,")
			loConsulta.AppendLine("            Campos_Nomina.Nom_Cam                   AS Nom_Cam,")
			loConsulta.AppendLine("            Campos_Nomina.Uso                       AS Uso,")
			loConsulta.AppendLine("            (CASE WHEN Campos_Nomina.Status='A'")
			loConsulta.AppendLine("                THEN 'Activo' ELSE 'Inactivo'")
			loConsulta.AppendLine("            END)                                    AS Estatus_Campo,")
			loConsulta.AppendLine("            Grupos_Conceptos.Cod_Gru                AS Cod_Gru,")
			loConsulta.AppendLine("            Grupos_Conceptos.nom_gru                AS Nom_Gru,")
			loConsulta.AppendLine("            (CASE COALESCE(Grupos_Conceptos.Status, '')")
			loConsulta.AppendLine("                WHEN 'A' THEN 'Activo'")
			loConsulta.AppendLine("                WHEN '' THEN NULL")
			loConsulta.AppendLine("                ELSE 'Inactivo'")
			loConsulta.AppendLine("            END)                                     AS Estatus_Grupo ")
			loConsulta.AppendLine("FROM        Campos_Nomina")
			loConsulta.AppendLine("    LEFT JOIN Renglones_Grupos_Campos")
			loConsulta.AppendLine("        ON  Renglones_Grupos_Campos.cod_cam = Campos_Nomina.cod_cam")
			loConsulta.AppendLine("    LEFT JOIN Grupos_Conceptos")
			loConsulta.AppendLine("        ON  Grupos_Conceptos.Cod_Gru = Renglones_Grupos_Campos.cod_gru")
			loConsulta.AppendLine("        AND Grupos_Conceptos.Cod_Gru BETWEEN " & lcParametro4Desde )
			loConsulta.AppendLine("            AND " & lcParametro4Hasta )
			loConsulta.AppendLine("WHERE       Campos_Nomina.Cod_Cam BETWEEN " & lcParametro0Desde )
			loConsulta.AppendLine("            AND " & lcParametro0Hasta )
			loConsulta.AppendLine("        AND Campos_Nomina.Uso IN (" & lcParametro1Desde & ")" )
			loConsulta.AppendLine("        AND Campos_Nomina.Status IN (" & lcParametro2Desde & ")" )
            If Not llIncluirCamposSinAsignar Then 
			    loConsulta.AppendLine("        AND Grupos_Conceptos.Cod_Gru IS NOT NULL" )
            End If 
			loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento )
			loConsulta.AppendLine("")


            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos()
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCamposNomina_ConSusGrupos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrCamposNomina_ConSusGrupos.ReportSource = loObjetoReporte
			
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
' RJG: 29/11/13: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
