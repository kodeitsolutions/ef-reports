Imports System.Data
Partial Class fAjustes_Precios_Ingles
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ajustes_Precios.Documento                           AS  Documento, ")
            loComandoSeleccionar.AppendLine("           Ajustes_Precios.Fec_Ini                             AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ajustes_Precios.Status                              AS  Status, ")
            loComandoSeleccionar.AppendLine("           Ajustes_Precios.Comentario                          AS  Observacion, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Renglon                          AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Cod_Art                          AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Pre_Ant                          AS  Pre_Ant, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Pre_Nue                          AS  Pre_Nue, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Cos_Pro1                         AS  Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Cos_Pro2                         AS  Cos_Pro2, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Cos_Ult1                         AS  Cos_Ult1, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Cos_Ult2                         AS  Cos_Ult2, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios.Tip_Pre                          AS  Tip_Pre, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art				                    AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Modelo                                    AS  Modelo, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Uni1                                  AS  Unidad ")
            loComandoSeleccionar.AppendLine(" FROM      Ajustes_Precios, ")
            loComandoSeleccionar.AppendLine("           Renglones_APrecios, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Ajustes_Precios.Documento       =   Renglones_APrecios.Documento ")
            loComandoSeleccionar.AppendLine("           And Renglones_APrecios.Cod_Art  =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Ajustes_Precios.Documento, Renglones_APrecios.Renglon ")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAjustes_Precios_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfAjustes_Precios_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 07/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'
