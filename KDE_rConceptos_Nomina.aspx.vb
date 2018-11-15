'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_rConceptos_Nomina"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_rConceptos_Nomina
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	    Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		    Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(4) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(4) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro2Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro2Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Recibos.Documento			AS Documento,")
            loConsulta.AppendLine("		    Recibos.Cod_Tra				AS Cod_Tra,")
            loConsulta.AppendLine("		    Trabajadores.Nom_Tra		AS Nom_Tra,")
            loConsulta.AppendLine("		    Recibos.Cod_Con				AS Contrato,")
            loConsulta.AppendLine("		    Renglones_Recibos.Cod_Con	AS Cod_Con,")
            loConsulta.AppendLine("		    Conceptos_Nomina.Nom_Con	AS Nom_Con,")
            loConsulta.AppendLine("		    Renglones_Recibos.Mon_Net	AS Mon_Net	")
            loConsulta.AppendLine("FROM Recibos")
            loConsulta.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            loConsulta.AppendLine("WHERE Recibos.Fecha BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("	AND Renglones_Recibos.Cod_Con BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")

            '      loConsulta.AppendLine("")
            '      loConsulta.AppendLine("SELECT   Cod_Con,")
            'loConsulta.AppendLine("         Nom_Con,")
            'loConsulta.AppendLine("         Status,")
            'loConsulta.AppendLine("         (CASE WHEN Status = 'A'")
            'loConsulta.AppendLine("             THEN 'Activo' ELSE 'Inactivo'")
            'loConsulta.AppendLine("          END) AS Estatus_Conceptos")
            'loConsulta.AppendLine("FROM     Conceptos_Nomina")
            'loConsulta.AppendLine("WHERE    Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde )
            'loConsulta.AppendLine("         AND " & lcParametro0Hasta )
            'loConsulta.AppendLine("     AND Conceptos_Nomina.Status IN (" & lcParametro1Desde & ")")
            'loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            'loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

           
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("KDE_rConceptos_Nomina", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvKDE_rConceptos_Nomina.ReportSource = loObjetoReporte
			  
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
