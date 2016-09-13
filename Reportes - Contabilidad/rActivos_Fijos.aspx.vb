Imports System.Data
Partial Class rActivos_Fijos
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()



            'lcComandoSelect = "SELECT	Cod_Act, " _
            '       & "Nom_Act, " _
            '       & "Status, " _
            '       & "(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Activos_Fijos " _
            '    & "FROM	 Activos_Fijos " _
            '    & "WHERE Cod_Act between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
            '       & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
            '       & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
            '    & " ORDER BY Cod_Act, " _
            '       & " Nom_Act "

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Activos_Fijos.Cod_Act,")
            loConsulta.AppendLine("            Activos_Fijos.Nom_Act,")
            loConsulta.AppendLine("			(CASE Activos_Fijos.Status ")
            loConsulta.AppendLine("				WHEN 'A' THEN	'Activo' ")
            loConsulta.AppendLine("				WHEN 'S' THEN	'Suspendido' ")
            loConsulta.AppendLine("				ELSE			'Inactivo' ")
            loConsulta.AppendLine("			END) AS status, ")
            loConsulta.AppendLine("            Activos_Fijos.seriales,")
            loConsulta.AppendLine("            Grupos_Activos.nom_gru,")
            loConsulta.AppendLine("            Ubicaciones_Activos.nom_ubi,")
            loConsulta.AppendLine("            Tipos_Activos.nom_tip,")
            loConsulta.AppendLine("            Centros_Costos.nom_cen")
            loConsulta.AppendLine("FROM        Activos_Fijos")
            loConsulta.AppendLine("    JOIN    Grupos_Activos ON Grupos_Activos.Cod_gru = Activos_Fijos.Cod_gru")
            loConsulta.AppendLine("    JOIN    Ubicaciones_Activos ON Ubicaciones_Activos.Cod_ubi = Activos_Fijos.Cod_ubi")
            loConsulta.AppendLine("    JOIN    Tipos_Activos ON Tipos_Activos.Cod_tip = Activos_Fijos.Cod_tip")
            loConsulta.AppendLine("    JOIN    Centros_Costos ON Centros_Costos.Cod_cen = Activos_Fijos.Cod_cen")
            loConsulta.AppendLine("WHERE	   Activos_Fijos.Cod_Act BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND Activos_Fijos.status IN (" & lcParametro1Desde & ")")
            If lcParametro2Desde <> "''" Then
                loConsulta.AppendLine(" 	AND Activos_Fijos.seriales IN (" & lcParametro2Desde & ")")
            End If
            loConsulta.AppendLine(" 	    AND Activos_Fijos.Cod_gru BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Activos_Fijos.Cod_ubi BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND Activos_Fijos.Cod_tip BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Activos_Fijos.Cod_cen BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("ORDER BY  " & lcOrdenamiento)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rActivos_Fijos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrActivos_Fijos.ReportSource = loObjetoReporte


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
' EAG:  07/08/15: Normalizar y agregar nuevos parametros
'-------------------------------------------------------------------------------------------'
