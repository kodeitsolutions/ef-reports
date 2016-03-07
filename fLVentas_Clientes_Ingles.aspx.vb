Imports System.Data
Partial Class fLVentas_Clientes_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Ventas.Cod_Cli, ")
            ''loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")

            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Libres_Ventas.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (CAST (Libres_Ventas.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE Libres_Ventas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Libres_Ventas.Nom_Cli AS VARCHAR) = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Libres_Ventas.Rif = '') THEN Clientes.Rif ELSE Libres_Ventas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Libres_Ventas.Nom_Cli AS VARCHAR) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Libres_Ventas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Libres_Ventas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Libres_Ventas.Nom_Cli AS VARCHAR) = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Libres_Ventas.Telefonos = '') THEN Clientes.Telefonos ELSE Libres_Ventas.Telefonos END) END) AS  Telefonos, ")

            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Ventas, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Ventas.Documento =   Renglones_LVentas.Documento AND ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Cli   =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   Renglones_LVentas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLVentas_Clientes_Ingles", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLVentas_Clientes_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 06/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 24/03/2011: Se cambio la etiqueta Client por Customer en el archivo rpt.
'-------------------------------------------------------------------------------------------'
