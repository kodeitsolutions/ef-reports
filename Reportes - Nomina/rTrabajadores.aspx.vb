﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrabajadores"
'-------------------------------------------------------------------------------------------'
Partial Class rTrabajadores
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("SELECT	Cod_Tra, ")
            loConsulta.AppendLine("         Nom_Tra, ")
            loConsulta.AppendLine("         Status, ")
            loConsulta.AppendLine("         (CASE WHEN Status = 'A' THEN 'Activo' ELSE 'Inactivo' END) AS Status_Trabajadores ")
            loConsulta.AppendLine("FROM	    Trabajadores ")
            loConsulta.AppendLine("WHERE    Cod_Tra BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         AND " & lcParametro0Hasta)
            loConsulta.AppendLine("         AND Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loConsulta.AppendLine(" ORDER BY Cod_Tra, Nom_Tra")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrabajadores.ReportSource = loObjetoReporte

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
' MJP: 09/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08 : Agragacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS: 20/05/09: Estandarizacion del codigo, combo del status, ordenamiento
'-------------------------------------------------------------------------------------------'
' RJG: 13/08/14: Ajuste del layout.                                                        '
'-------------------------------------------------------------------------------------------'
