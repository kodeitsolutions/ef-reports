'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fLibres_Rma"
'-------------------------------------------------------------------------------------------'
Partial Class fLibres_Rma

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	libres_rma.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Documento, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           libres_rma.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Renglon, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Precio1, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      libres_rma, ")
            loComandoSeleccionar.AppendLine("           renglones_lrma, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     libres_rma.Documento =   renglones_lrma.Documento AND ")
            loComandoSeleccionar.AppendLine("           libres_rma.Cod_Cli   =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           libres_rma.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           libres_rma.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   renglones_lrma.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")
            
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLibres_Rma", laDatosReporte)

			CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLibres_Rma.ReportSource = loObjetoReporte

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
' CMS: 20/02/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
