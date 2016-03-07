'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rVendedores
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()
			loConsulta.AppendLine("SELECT		Vendedores.Cod_Ven,    " )
			loConsulta.AppendLine("				Vendedores.Nom_Ven,    " )
			loConsulta.AppendLine("				Vendedores.Status,     " )
			loConsulta.AppendLine("				(CASE WHEN Status = 'A' " )
			loConsulta.AppendLine("				    THEN 'Activo' ELSE 'Inactivo' " )
			loConsulta.AppendLine("				END) AS Status_Vendedores" )
			loConsulta.AppendLine("FROM			Vendedores " )
			loConsulta.AppendLine("WHERE		Cod_Ven BETWEEN " & lcParametro0Desde)
			loConsulta.AppendLine("				AND " & lcParametro0Hasta)
			loConsulta.AppendLine("			AND Status IN (" & lcParametro1Desde & ")")
			loConsulta.AppendLine("ORDER BY     " & lcOrdenamiento)

			
			Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
 
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVendedores", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrVendedores.ReportSource = loObjetoReporte
			
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
' MJP: 07/07/08: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08: Creación objeto que cierra el archivo de reporte. Agregacion filtro Status.'
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.                 '
'-------------------------------------------------------------------------------------------'
' GCR: 02/01/09: Estandarizacion de código y ajustes al diseño.                             '
'-------------------------------------------------------------------------------------------'
' RJG: 13/06/14: Estandarizacion de código y ajustes al diseño.                             '
'-------------------------------------------------------------------------------------------'
