'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCasos_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rCasos_Articulos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
		   
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
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
			 

            loConsulta.AppendLine("SELECT	Articulos.Cod_Art            AS Cod_Reg,")
            loConsulta.AppendLine(" 		Articulos.Nom_Art            AS Nom_Reg,")
            loConsulta.AppendLine(" 		Casos.Status                AS Status,")
            loConsulta.AppendLine(" 		Casos.Documento             AS Documento,")
            loConsulta.AppendLine(" 		Casos.Fec_ini               AS Fec_ini,")
            loConsulta.AppendLine(" 		Casos.Fec_fin               AS Fec_fin,")
            loConsulta.AppendLine(" 		(CASE Casos.Status")
            loConsulta.AppendLine(" 		    WHEN 'Pendiente' ")
            loConsulta.AppendLine(" 		    THEN DATEDIFF(DAY,Casos.Fec_ini, GETDATE())")
            loConsulta.AppendLine(" 		    ELSE DATEDIFF(DAY,Casos.Fec_ini, Casos.Fec_fin)")
            loConsulta.AppendLine(" 		END)                        AS Dias,")
            loConsulta.AppendLine(" 		Casos.Fec_fin               AS Fec_fin,")
            loConsulta.AppendLine(" 		Casos.Cod_Con               AS Cod_Con,")
            loConsulta.AppendLine(" 		Casos.Cod_Coo               AS Cod_Coo,")
            loConsulta.AppendLine(" 		Casos.Cod_Eje               AS Cod_Eje,")
            loConsulta.AppendLine(" 		Casos.Control               AS Control,")
            loConsulta.AppendLine(" 		Casos.Asunto                AS Asunto,")
            loConsulta.AppendLine(" 		Casos.Comentario            AS Comentario")
            loConsulta.AppendLine("FROM   Casos")
            loConsulta.AppendLine("    LEFT JOIN Articulos ON Articulos.Cod_Art = Casos.Cod_Art")
            loConsulta.AppendLine("WHERE      Casos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Status	IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine(" 	    AND Casos.Cod_Reg	BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Coo	BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Eje   BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Casos.Cod_Suc	BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("ORDER BY Articulos.Cod_Art, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCasos_Articulos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCasos_Articulos.ReportSource = loObjetoReporte


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
' RJG: 25/06/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
