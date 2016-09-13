Imports System.Data
Partial Class rGrupos_Activos
   Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            'lcComandoSelect = "SELECT	Cod_Gru, " _
            '       & "Nom_Gru, " _
            '       & "Status, " _
            '       & "(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Grupos_Activos " _
            '    & "FROM	 Grupos_Activos " _
            '    & "WHERE Cod_Gru between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
            '       & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
            '    & " ORDER BY Cod_Gru, " _
            '       & " Nom_Gru "

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Cod_Gru,")
            loConsulta.AppendLine("         Nom_Gru,")
            loConsulta.AppendLine("			(CASE Status ")
            loConsulta.AppendLine("				WHEN 'A' THEN	'Activo' ")
            loConsulta.AppendLine("				WHEN 'S' THEN	'Suspendido' ")
            loConsulta.AppendLine("				ELSE			'Inactivo' ")
            loConsulta.AppendLine("			END) AS status ")
            loConsulta.AppendLine("FROM        Grupos_Activos")
            loConsulta.AppendLine("WHERE	   Cod_Gru BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine("ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGrupos_Activos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrGrupos_Activos.ReportSource = loObjetoReporte


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
' MJP   :  15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------' 
' MVP:  05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' EAG:  07/08/15: Acomodar parametro Estatus estaba como caracter, se pasó a lista
'-------------------------------------------------------------------------------------------'
