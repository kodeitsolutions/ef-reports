Imports System.Data
Partial Class fPresupuestos_Proveedores_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Presupuestos.Cod_Pro                    As  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro                     As  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Documento, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
            loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("			ELSE Renglones_Presupuestos.Notas")
            loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Mon_Imp1         As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Por_Des         As  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Mon_Des         As  Descuento ")
            'loComandoSeleccionar.AppendLine(" FROM      Presupuestos, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Presupuestos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores, ")
            'loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            'loComandoSeleccionar.AppendLine("           Vendedores, ")
            'loComandoSeleccionar.AppendLine("           Articulos ")           
            loComandoSeleccionar.AppendLine(" FROM      Presupuestos ")
            loComandoSeleccionar.AppendLine("           JOIN Renglones_Presupuestos on Presupuestos.Documento  =   Renglones_Presupuestos.Documento")
            loComandoSeleccionar.AppendLine("           JOIN Proveedores ON Presupuestos.Cod_Pro    =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           JOIN Formas_Pagos ON Presupuestos.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Vendedores ON Presupuestos.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           JOIN Articulos ON Articulos.Cod_Art       =   Renglones_Presupuestos.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(, ", "(")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPresupuestos_Proveedores_Ingles", laDatosReporte)

            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPresupuestos_Proveedores_Ingles.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                      "No se pudo Completar el Proceso: " & loExcepcion.Message & " StackTrace: " & loExcepcion.StackTrace, _
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
' DLC: 08/07/2010: Se agrego a la consulta, el porcentaje de descuento y porcentaje de recargo
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/2011: Se hicieron ajustes para impresion en el arcivo rpt.
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select, Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'
