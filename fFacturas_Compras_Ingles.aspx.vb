﻿Imports System.Data
Partial Class fFacturas_Compras_Ingles
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Rif = '') THEN Proveedores.Rif ELSE Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Compras.Telefonos END) END) AS  Telefonos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Compras.Nom_Pro As Nombre_Generico, ")
            loComandoSeleccionar.AppendLine("           Compras.Rif As Rif_Generico, ")
            loComandoSeleccionar.AppendLine("           Compras.Nit As Nit_Generico, ")
            loComandoSeleccionar.AppendLine("           Compras.Dir_Fis As Dir_Fis_Generico, ")
            loComandoSeleccionar.AppendLine("           Compras.Telefonos As Telefonos_Generico, ")
            loComandoSeleccionar.AppendLine("           Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Compras.Factura, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Compras.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Compras.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,40)   AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
            loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("			ELSE Renglones_Compras.Notas")
            loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Uni1          As Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Mon_Net   As Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Por_Imp1  As Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Mon_Imp1  As Impuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Por_Des  As Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Mon_Des  As Descuento ")
            loComandoSeleccionar.AppendLine(" FROM      Compras, ")
            loComandoSeleccionar.AppendLine("           Renglones_Compras, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Compras.Documento   =   Renglones_Compras.Documento AND ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Pro     =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_For     =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Ven     =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Compras.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim lcImpuesto As String
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Compras_Ingles", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Compras_Ingles.ReportSource = loObjetoReporte

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
' DLC: 10/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 28/06/2010: Ajuste del caracter de decimal y caracter de miles en los 
'                               porcentaje impuesto.
'-------------------------------------------------------------------------------------------'
' DLC: 08/07/2010: Se Agrego la columna de porcentaje de descuento y monto de descuento.
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
