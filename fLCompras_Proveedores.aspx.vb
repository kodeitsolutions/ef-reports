Imports System.Data
Partial Class fLCompras_Proveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Libres_Compras.Nom_Pro As Varchar) = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cast(Libres_Compras.Nom_Pro As Varchar)= '') THEN Proveedores.Nom_Pro ELSE Libres_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Libres_Compras.Nom_Pro As Varchar) = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Libres_Compras.Rif = '') THEN Proveedores.Rif ELSE Libres_Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Libres_Compras.Nom_Pro As Varchar)= '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Libres_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Libres_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Libres_Compras.Nom_Pro As Varchar) = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Libres_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Libres_Compras.Telefonos END) END) AS  Telefonos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Por_Des As Por_Des1_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Compras, ")
            loComandoSeleccionar.AppendLine("           Renglones_LCompras, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Compras.Documento =   Renglones_LCompras.Documento AND ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Cod_Pro   =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Libres_Compras.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art        =   Renglones_LCompras.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLCompras_Proveedores", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text35"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLCompras_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 24/12/09: Se le incluyo los montos de descuentos y recargos.
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Se coloco la validación de registros 0 y el metodo de carga de imagen 
'					Se agrego la distribucion de impuesto y el descuento del renglon
'-------------------------------------------------------------------------------------------'