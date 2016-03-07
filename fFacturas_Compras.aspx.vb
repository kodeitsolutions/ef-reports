'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Compras.Cod_Pro, ")
            loConsulta.AppendLine("       (CASE WHEN (Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Compras.Nom_Pro END) AS  Nom_Pro, ")
            loConsulta.AppendLine("       (CASE WHEN (Compras.Rif = '') THEN Proveedores.Rif ELSE Compras.Rif END) AS  Rif, ")
            loConsulta.AppendLine("       Proveedores.Nit, ")
            loConsulta.AppendLine("       (CASE WHEN (Compras.Dir_Fis = '') THEN Proveedores.Dir_Fis ELSE Compras.Dir_Fis END) AS  Dir_Fis, ")
            loConsulta.AppendLine("       (CASE WHEN (Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Compras.Telefonos END) AS  Telefonos, ")
            loConsulta.AppendLine("       Proveedores.Fax, ")
            loConsulta.AppendLine("       Compras.Documento, ")
            loConsulta.AppendLine("       Compras.Factura, ")
            loConsulta.AppendLine("       Compras.Control, ")
            loConsulta.AppendLine("       Compras.Fec_Ini, ")
            loConsulta.AppendLine("       Compras.Fec_Fin, ")
            loConsulta.AppendLine("       Compras.Nom_Pro   AS Nombre_Generico, ")
            loConsulta.AppendLine("       Compras.Rif       AS Rif_Generico, ")
            loConsulta.AppendLine("       Compras.Nit       AS Nit_Generico, ")
            loConsulta.AppendLine("       Compras.Dir_Fis   AS Dir_Fis_Generico, ")
            loConsulta.AppendLine("       Compras.Telefonos AS Telefonos_Generico,")
            loConsulta.AppendLine("       Compras.Mon_Bru, ")
            loConsulta.AppendLine("       Compras.Mon_Imp1, ")
            loConsulta.AppendLine("       Compras.Por_Des1, ")
            loConsulta.AppendLine("       Compras.Mon_Des1, ")
            loConsulta.AppendLine("       Compras.Por_Rec1, ")
            loConsulta.AppendLine("       Compras.Mon_Rec1, ")
            loConsulta.AppendLine("       Compras.Dis_Imp, ")
            loConsulta.AppendLine("       Compras.Mon_Net, ")
            loConsulta.AppendLine("       Compras.Cod_For, ")
            loConsulta.AppendLine("       Compras.Cod_Mon, ")
            loConsulta.AppendLine("       Compras.Por_Imp1, ")
            loConsulta.AppendLine("       Compras.Comentario, ")
            loConsulta.AppendLine("       Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("       Compras.Cod_Ven, ")
            loConsulta.AppendLine("       Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("       Renglones_Compras.Cod_Art, ")
            loConsulta.AppendLine("       (CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ELSE Renglones_Compras.Notas END) AS Nom_Art,  ")
            loConsulta.AppendLine("       Renglones_Compras.Renglon, ")
            loConsulta.AppendLine("       Renglones_Compras.Can_Art1, ")
            loConsulta.AppendLine("       Articulos.Cod_Uni1          As Cod_Uni, ")
            loConsulta.AppendLine("       Renglones_Compras.Precio1, ")
            loConsulta.AppendLine("       Renglones_Compras.Mon_Net   As Neto, ")
            loConsulta.AppendLine("       Renglones_Compras.Cod_Imp, ")
            loConsulta.AppendLine("       Renglones_Compras.Por_Des  As Por_Des, ")
            loConsulta.AppendLine("       Renglones_Compras.Por_Imp1  As Por_Imp, ")
            loConsulta.AppendLine("       Renglones_Compras.Mon_Imp1  As Impuesto ")
            loConsulta.AppendLine("FROM    Compras")
            loConsulta.AppendLine("   JOIN Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            loConsulta.AppendLine("   JOIN Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loConsulta.AppendLine("   JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Compras.Cod_For")
            loConsulta.AppendLine("   JOIN Vendedores ON  Vendedores.Cod_Ven = Compras.Cod_Ven")
            loConsulta.AppendLine("   JOIN Articulos ON Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loConsulta.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                   ' If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                   ' End If
                Next loImpuesto
            Next lnNumeroFila

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Compras", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Compras.ReportSource = loObjetoReporte

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
' JJD: 13/11/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 13/11/09: Se cambio el campo Cos_Ult1 por Precio1
'-------------------------------------------------------------------------------------------'
' CMS: 04/03/10: Se coloco la validación de registros 0 y el metodo de carga de imagen 
'-------------------------------------------------------------------------------------------'
' CMS: 03/08/10: Se coloco el porcentaje de descuento del renglon
'-------------------------------------------------------------------------------------------'
' MAT: 10/05/11: Adición de nuevos campos: Factura y Control
'-------------------------------------------------------------------------------------------'
' MAT: 10/05/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT: 14/03/15: Simplificación del SELECT 
'-------------------------------------------------------------------------------------------'
