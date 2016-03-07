'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fEstimaciones"
'-------------------------------------------------------------------------------------------'
Partial Class fEstimaciones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Estimaciones.Documento, ")
            loConsulta.AppendLine("        Estimaciones.Control, ")
            loConsulta.AppendLine("        Estimaciones.Referencia, ")
            loConsulta.AppendLine("        Estimaciones.Fec_Ini, ")
            loConsulta.AppendLine("        Estimaciones.Fec_Fin, ")
            loConsulta.AppendLine("        Estimaciones.Mon_Bru, ")
            loConsulta.AppendLine("        Estimaciones.Mon_Imp1, ")
            loConsulta.AppendLine("        Estimaciones.Por_Des1, ")
            loConsulta.AppendLine("        Estimaciones.Mon_Des1, ")
            loConsulta.AppendLine("        Estimaciones.Por_Rec1, ")
            loConsulta.AppendLine("        Estimaciones.Mon_Rec1, ")
            loConsulta.AppendLine("        Estimaciones.Dis_Imp, ")
            loConsulta.AppendLine("        Estimaciones.Mon_Net, ")
            loConsulta.AppendLine("        Estimaciones.Cod_Mon, ")
            loConsulta.AppendLine("        Estimaciones.Por_Imp1, ")
            loConsulta.AppendLine("        Estimaciones.Comentario, ")
            loConsulta.AppendLine("        Estimaciones.Tip_Ori, ")
            loConsulta.AppendLine("        Estimaciones.Cla_Ori, ")
            loConsulta.AppendLine("        Estimaciones.Doc_Ori, ")
            loConsulta.AppendLine("        Estimaciones.Origen, ")
            loConsulta.AppendLine("        Estimaciones.Adicional, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Cod_Art, ")
            loConsulta.AppendLine("        (CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ELSE Renglones_Estimaciones.Nom_Art END) AS Nom_Art,  ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Renglon, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Can_Art1, ")
            loConsulta.AppendLine("        Articulos.Cod_Uni1          As Cod_Uni, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Precio1, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Mon_Net   As Neto, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Cod_Imp, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Por_Imp1  As Por_Imp, ")
            loConsulta.AppendLine("        Renglones_Estimaciones.Mon_Imp1  As Impuesto ")
            loConsulta.AppendLine("FROM    Estimaciones")
            loConsulta.AppendLine("    JOIN Renglones_Estimaciones ON Renglones_Estimaciones.Documento = Estimaciones.Documento")
            loConsulta.AppendLine("    JOIN Articulos ON Renglones_Estimaciones.Cod_Art = Articulos.Cod_Art")
            loConsulta.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Solo la primera fila
            If (laDatosReporte.Tables(0).Rows.Count > 0) Then
                lcXml = laDatosReporte.Tables(0).Rows(0).Item("dis_imp")

                If Not String.IsNullOrEmpty(lcXml.Trim()) Then

                    loImpuestos.LoadXml(lcXml)

                    'En cada renglón lee el contenido de la distribució de impuestos
                    For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                       ' If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                            lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                       ' End If
                    Next loImpuesto

                End If

            End If

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fEstimaciones", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfEstimaciones.ReportSource = loObjetoReporte

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
' RJG: 14/03/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
