Imports System.Data
Partial Class fCortes_Inventarios

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cortes.Documento                            AS  Documento, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cortes.Fec_Ini, 103)     AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cortes.Cod_Alm                              AS  Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Cortes.Status                               AS  Status, ")
            loComandoSeleccionar.AppendLine("           Cortes.Comentario                           AS  Observacion, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Renglon                    AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Cod_Art                    AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Can_Teo                    AS  Can_Teo, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Can_Rea                    AS  Can_Rea, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Costo                      AS  Costo, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Cos_Pro1                   AS  Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Cos_Pro2                   AS  Cos_Pro2, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Cos_Ult1                   AS  Cos_Ult1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes.Cos_Ult2                   AS  Cos_Ult2, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art                           AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Modelo                            AS  Modelo, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Uni1                          AS  Unidad ")
            loComandoSeleccionar.AppendLine(" FROM      Cortes, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cortes, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Cortes.Documento                =   Renglones_Cortes.Documento ")
            loComandoSeleccionar.AppendLine("           And Renglones_Cortes.Cod_Art    =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Cortes.Documento, Renglones_Cortes.Renglon ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCortes_Inventarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCortes_Inventarios.ReportSource = loObjetoReporte

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
' JJD: 17/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/10: Se coloco la validación de registros 0 y el metodo de carga de imagen 
'-------------------------------------------------------------------------------------------'