Imports System.Data
Partial Class fRequisiciones_Internas_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Requisiciones.Cod_Pro                    As  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro                     As  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           SPACE(1)                                 As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           SPACE(1)                                 As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,24)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Mon_Imp1         As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Por_Des         As  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones.Mon_Des         As  Descuento ")
            loComandoSeleccionar.AppendLine(" FROM      Requisiciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Requisiciones.Documento  =   Renglones_Requisiciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Pro    =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art      =   Renglones_Requisiciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(, ", "(")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fRequisiciones_Internas_Ingles", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfRequisiciones_Internas_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 10/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/2011: Se hicieron los ajustes para Impresion.
'-------------------------------------------------------------------------------------------'
' MAT: 14/03/11: Corrección de las unidades del reporte según requerimientos
'-------------------------------------------------------------------------------------------'
' MAT: 05/09/11: Creación de los parámetros para las leyendas en el formato
'-------------------------------------------------------------------------------------------'

