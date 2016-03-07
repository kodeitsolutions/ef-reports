'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIndicador_Requisiciones_RecVsProc"
'-------------------------------------------------------------------------------------------'
Partial Class rIndicador_Requisiciones_RecVsProc
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Vendedores.Cod_Ven                      AS Cod_Ven,")
            loConsulta.AppendLine("        Vendedores.Nom_Ven                      AS Nom_Ven,")
            loConsulta.AppendLine("        COUNT(*)                                AS Recibidos,")
            loConsulta.AppendLine("        SUM(CASE Requisiciones.Status")
            loConsulta.AppendLine("            WHEN 'Procesado' THEN 1 ELSE 0")
            loConsulta.AppendLine("        END)                                    AS Procesados,")
            loConsulta.AppendLine("        CAST(SUM(CASE Requisiciones.Status")
            loConsulta.AppendLine("            WHEN 'Procesado' THEN 1 ELSE 0")
            loConsulta.AppendLine("        END) AS DECIMAL(28,10))*100")
            loConsulta.AppendLine("        /CAST(COUNT(*) AS DECIMAL(28,10))       AS Porcentaje")
            loConsulta.AppendLine("FROM    Requisiciones")
            loConsulta.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Requisiciones.Cod_Ven")
            loConsulta.AppendLine("WHERE   Requisiciones.Status IN ('Confirmado' , 'Afectado', 'Procesado')")
            loConsulta.AppendLine("    AND Requisiciones.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Requisiciones.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND Requisiciones.Cod_Pro BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Requisiciones.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         AND " & lcParametro3Hasta)
            loConsulta.AppendLine("    AND Requisiciones.Cod_Mon BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("         AND " & lcParametro4Hasta)
            If lcParametro5Desde.ToUpper().Trim() = "IGUAL" Then
                loConsulta.AppendLine("    AND Requisiciones.Cod_Rev BETWEEN " & lcParametro6Desde)
                loConsulta.AppendLine("         AND " & lcParametro6Hasta)
            Else
                loConsulta.AppendLine("    AND Requisiciones.Cod_Rev NOT BETWEEN " & lcParametro6Desde)
                loConsulta.AppendLine("         AND " & lcParametro6Hasta)
            End If
            loConsulta.AppendLine("    AND Requisiciones.Cod_Suc BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("         AND " & lcParametro7Hasta)
            loConsulta.AppendLine("GROUP BY Vendedores.Cod_Ven, Vendedores.Nom_Ven")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIndicador_Requisiciones_RecVsProc", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrIndicador_Requisiciones_RecVsProc.ReportSource = loObjetoReporte

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
' RJG: 15/11/14: Codigo inicial
'-------------------------------------------------------------------------------------------'
