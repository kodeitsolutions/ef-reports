'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosArticulos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fFormulas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Formulas.Documento                                          AS Documento,")
            loComandoSeleccionar.AppendLine("           Formulas.Fec_Ini                                            AS Fecha,")
            loComandoSeleccionar.AppendLine("           Formulas.Comentario                                         AS Observacion,")
            loComandoSeleccionar.AppendLine("           Formulas.Cod_Art                                            AS Cod_Art,")
            loComandoSeleccionar.AppendLine("           Formulas.Referencia                                         AS Referencia,")
            loComandoSeleccionar.AppendLine("           Formulas.Mon_Net                                            AS Mon_Net,")
            loComandoSeleccionar.AppendLine("           Formulas.Cos_Uni                                            AS Costo_Uni,")
            loComandoSeleccionar.AppendLine("           Formulas.Caracter1                                          AS Lote,")
            loComandoSeleccionar.AppendLine("           Formulas.Clase                                              AS Clase,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Renglon                                  AS Renglon,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Cod_Art                                  AS Codigo,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Can_Art1                                 AS Cantidad,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Cod_Uni                                  AS Unidad,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Cos_Ult1                                 AS Costo,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Notas                                    AS Notas,")
            loComandoSeleccionar.AppendLine("           (Renglones_Formulas.Can_Art1 * Renglones_Formulas.Cos_Ult1) AS Total,")
            loComandoSeleccionar.AppendLine("           Renglones_Formulas.Notas                                    AS Descripcion,")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art                                           AS Nom_Art")
            loComandoSeleccionar.AppendLine(" FROM      Formulas")
            loComandoSeleccionar.AppendLine("       JOIN Renglones_Formulas ON Formulas.Documento = Renglones_Formulas.Documento ")
            loComandoSeleccionar.AppendLine("       JOIN Articulos ON Formulas.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fFormulas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fFormulas.ReportSource = loObjetoReporte

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
' GMO: 10/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Adecuacion a la estructura de los reportes eFactory.
'-------------------------------------------------------------------------------------------'
' CMS: 28/06/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'