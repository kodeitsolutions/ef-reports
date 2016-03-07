'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAuditorias_Requisiciones"
'-------------------------------------------------------------------------------------------'
Partial Class fAuditorias_Requisiciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Auditorias.Documento, ")
			loConsulta.AppendLine("            Auditorias.Cod_Usu,")
			loConsulta.AppendLine("            CONVERT(NCHAR(10), Auditorias.Registro, 103)            AS Fecha, ")
			loConsulta.AppendLine("            LEFT(CONVERT(VARCHAR(30), Auditorias.Registro, 108), 5) AS Hora,")
			loConsulta.AppendLine("            Auditorias.Tipo, ")
			loConsulta.AppendLine("            Auditorias.Tabla, ")
			loConsulta.AppendLine("            Auditorias.Cod_Obj, ")
			loConsulta.AppendLine("            Auditorias.Notas, ")
			loConsulta.AppendLine("            Auditorias.Opcion, ")
			loConsulta.AppendLine("            Auditorias.Accion, ")
			loConsulta.AppendLine("            Auditorias.Equipo, ")
			loConsulta.AppendLine("            Auditorias.Detalle")
			loConsulta.AppendLine("FROM        Requisiciones")
			loConsulta.AppendLine("    JOIN    Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Requisiciones.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Requisiciones' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('RequisicionesInternas', 'Sin opción')")
            loConsulta.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loConsulta.AppendLine("ORDER BY    Auditorias.Registro DESC")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
		    
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

           '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
			'-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
						vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAuditorias_Requisiciones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvfAuditorias_Requisiciones.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 08/12/14: Codigo inicial, a partir de fAuditorias_Pedidos.                           '
'-------------------------------------------------------------------------------------------'
