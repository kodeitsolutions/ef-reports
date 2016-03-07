Imports System.Data
Partial Class fCortes_Inventarios_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cortes.Documento                            AS  Documento, ")
            loComandoSeleccionar.AppendLine("           Cortes.Fec_Ini                              AS  Fec_Ini, ")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCortes_Inventarios_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCortes_Inventarios_Ingles.ReportSource = loObjetoReporte

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
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
