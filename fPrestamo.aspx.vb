'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPrestamo"
'-------------------------------------------------------------------------------------------'
Partial Class fPrestamo
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Prestamos.Documento                      AS Documento,")
            loConsulta.AppendLine("        Prestamos.Fec_Asi                        AS Fec_Asi,")
            loConsulta.AppendLine("        Prestamos.Status                         AS Status,")
            loConsulta.AppendLine("        Prestamos.Cod_Tra                        AS Cod_Tra,")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra                     AS Nom_Tra,")
            loConsulta.AppendLine("        Trabajadores.Cedula                      AS Cedula,")
            loConsulta.AppendLine("        Prestamos.Cod_Tip                        AS Cod_Tip,")
            loConsulta.AppendLine("        Prestamos.Tip_Int                        AS Tip_Int,")
            loConsulta.AppendLine("        Prestamos.Por_Int                        AS Por_Int,")
            loConsulta.AppendLine("        Prestamos.Mon_Int                        AS Mon_Int,")
            loConsulta.AppendLine("        Prestamos.Mon_Bas                        AS Mon_Bas,")
            loConsulta.AppendLine("        Prestamos.Mon_Net                        AS Mon_Net,")
            loConsulta.AppendLine("        Prestamos.Mon_Sal                        AS Mon_Sal,")
            loConsulta.AppendLine("        Prestamos.Mon_Pag                        AS Mon_Pag,")
            loConsulta.AppendLine("        Prestamos.Mon_Abo                        AS Mon_Abo,")
            loConsulta.AppendLine("        Prestamos.Comentario                     AS Comentario,")
            loConsulta.AppendLine("        ROW_NUMBER() OVER(ORDER BY Fec_Ven ASC)  AS Renglon,  ")
            loConsulta.AppendLine("        Detalles_Prestamos.Fec_Ven               AS R_Fec_Ven,")
            loConsulta.AppendLine("        Detalles_Prestamos.Mon_Cap               AS R_Mon_Cap,") 
            loConsulta.AppendLine("        Detalles_Prestamos.Mon_Int               AS R_Mon_Int,") 
            loConsulta.AppendLine("        Detalles_Prestamos.Mon_Net               AS R_Mon_Net,") 
            loConsulta.AppendLine("        Detalles_Prestamos.Pagado                AS R_Pagado,") 
            loConsulta.AppendLine("        Detalles_Prestamos.Comentario            AS R_Comentario") 
            loConsulta.AppendLine("FROM    Prestamos")
            loConsulta.AppendLine("    JOIN Detalles_Prestamos ON Detalles_Prestamos.documento = Prestamos.Documento")
            loConsulta.AppendLine("JOIN	Trabajadores")
            loConsulta.AppendLine("	ON	Trabajadores.Cod_Tra = Prestamos.Cod_Tra ")
            loConsulta.AppendLine("WHERE	" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("ORDER BY fec_ven ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
           
			Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPrestamo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPrestamo.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 03/02/14: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
