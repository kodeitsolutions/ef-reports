'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCotizaciones_SalesT"
'-------------------------------------------------------------------------------------------'
Partial Class fCotizaciones_SalesT

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Comentario AS Comentario_renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Cotizaciones ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Cotizaciones ON (Cotizaciones.Documento  =   Renglones_Cotizaciones.Documento) ")
            loComandoSeleccionar.AppendLine(" JOIN  Clientes ON (Cotizaciones.Cod_Cli    =   Clientes.Cod_Cli)")
            loComandoSeleccionar.AppendLine(" JOIN  Formas_Pagos ON (Cotizaciones.Cod_For    =   Formas_Pagos.Cod_For)")
            loComandoSeleccionar.AppendLine(" JOIN  Vendedores ON (Cotizaciones.Cod_Ven    =   Vendedores.Cod_Ven)")
            loComandoSeleccionar.AppendLine(" JOIN  Articulos ON (Articulos.Cod_Art       =   Renglones_Cotizaciones.Cod_Art)")
            loComandoSeleccionar.AppendLine(" WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
            
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
								If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
									lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_SalesT", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text38"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCotizaciones_SalesT.ReportSource = loObjetoReporte

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
' MAT: 17/08/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
