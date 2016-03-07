Imports System.Data

Partial Class fDevoluciones_IHP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Nom_Pro        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Documento, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Por_Des  As  Por_Des_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Devoluciones_Proveedores, ")
            loComandoSeleccionar.AppendLine("           Renglones_DProveedores, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Devoluciones_Proveedores.Documento      =   Renglones_DProveedores.Documento AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_Pro        =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_For        =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_Ven        =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art                       =   Renglones_DProveedores.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuestos As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpuestos = "("

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
                        lcPorcentajesImpuestos = lcPorcentajesImpuestos & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuestos = lcPorcentajesImpuestos & ")"
            lcPorcentajesImpuestos = lcPorcentajesImpuestos.Replace("(,", "(")


            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDevoluciones_IHP", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuestos.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDevoluciones_IHP.ReportSource = loObjetoReporte


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
' MAT: 20/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
