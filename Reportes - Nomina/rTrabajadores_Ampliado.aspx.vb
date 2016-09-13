'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrabajadores_Ampliado"
'-------------------------------------------------------------------------------------------'
Partial Class rTrabajadores_Ampliado
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("SELECT  Trabajadores.Cod_Tra                        AS Cod_Tra,")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra                        AS Nom_Tra,")
            loConsulta.AppendLine("        Trabajadores.Fec_Ini                        AS Fec_Ini,")
            loConsulta.AppendLine("        COALESCE(Renglones_Campos_Nomina.Val_Num, 0)AS Sueldo,")
            loConsulta.AppendLine("        Trabajadores.Status                         AS Status,")
            loConsulta.AppendLine("        COALESCE(Cargos.nom_car,'')                 AS Nom_Car,")
            loConsulta.AppendLine("        (CASE Trabajadores.Status")
            loConsulta.AppendLine("            WHEN 'A' THEN 'Activo' ")
            loConsulta.AppendLine("            WHEN 'I' THEN 'Inactivo' ")
            loConsulta.AppendLine("            WHEN 'S' THEN 'Suspendido' ")
            loConsulta.AppendLine("            WHEN 'L' THEN 'Liquidado' ")
            loConsulta.AppendLine("            ELSE '[Desconocido]' ")
            loConsulta.AppendLine("        END)                                        AS Status_Trabajadores ")
            loConsulta.AppendLine("FROM	   Trabajadores")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina")
            loConsulta.AppendLine("        ON  Renglones_Campos_Nomina.Cod_tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Renglones_Campos_Nomina.Cod_Cam = 'A001'")
            loConsulta.AppendLine("    LEFT JOIN Cargos")
            loConsulta.AppendLine("        ON  Cargos.Cod_car = Trabajadores.Cod_car")
            loConsulta.AppendLine("WHERE   Trabajadores.Tip_Tra = 'Trabajador'")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("        AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("        AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("        AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("        AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("        AND " & lcParametro5Hasta)
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.toString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrabajadores_Ampliado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrabajadores_Ampliado.ReportSource = loObjetoReporte

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
' RJG: 13/08/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
