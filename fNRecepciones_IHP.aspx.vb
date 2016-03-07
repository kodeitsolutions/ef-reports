Imports System.Data
Partial Class fNRecepciones_IHP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Recepciones.Cod_Pro                    As  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Recepciones.Nom_Pro END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Rif = '') THEN Proveedores.Rif ELSE Recepciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Recepciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Recepciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Recepciones.Telefonos END) END) AS  Telefonos, ")
            
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Recepciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,25)    AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loComandoSeleccionar.AppendLine("			ELSE Renglones_Recepciones.Notas")
			loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")            
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Por_Des          As  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Recepciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Recepciones.Documento  =   Renglones_Recepciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Pro    =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art      =   Renglones_Recepciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
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
								If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
									lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
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


			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fNRecepciones_IHP", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")          
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfNRecepciones_IHP.ReportSource = loObjetoReporte

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
' MAT: 20/08/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
