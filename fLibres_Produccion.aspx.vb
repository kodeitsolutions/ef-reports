'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
										 
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fLibres_Produccion"
'-------------------------------------------------------------------------------------------'
Partial Class fLibres_Produccion

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Produccion.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Produccion, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Produccion.Documento =   Renglones_LProduccion.Documento AND ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Cli   =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   Renglones_LProduccion.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLibres_Produccion", laDatosReporte)

			CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString.Replace(".", ",")

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLibres_Produccion.ReportSource = loObjetoReporte

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
' CMS: 23/02/10: Codigo inicial 
'-------------------------------------------------------------------------------------------'
