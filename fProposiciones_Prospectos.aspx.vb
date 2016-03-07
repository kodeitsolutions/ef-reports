'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fProposiciones_Prospectos"
'-------------------------------------------------------------------------------------------'
Partial Class fProposiciones_Prospectos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Proposiciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Prospectos.Generico = 0 AND CAST (Proposiciones.Nom_Pro AS VARCHAR) = '') THEN Prospectos.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Proposiciones.Nom_Pro = '') THEN Prospectos.Nom_Pro ELSE Proposiciones.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Prospectos.Generico = 0 AND CAST (Proposiciones.Nom_Pro AS VARCHAR) = '') THEN Prospectos.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Proposiciones.Rif = '') THEN Prospectos.Rif ELSE Proposiciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Prospectos.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Prospectos.Generico = 0 AND CAST (Proposiciones.Nom_Pro AS VARCHAR) = '') THEN SUBSTRING(Prospectos.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Proposiciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Prospectos.Dir_Fis,1, 200) ELSE SUBSTRING(Proposiciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Prospectos.Generico = 0 AND CAST (Proposiciones.Nom_Pro AS VARCHAR) = '') THEN Prospectos.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Proposiciones.Telefonos = '') THEN Prospectos.Telefonos ELSE Proposiciones.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Prospectos.Fax, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Proposiciones.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Comentario AS Comentario_renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Proposiciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones, ")
            loComandoSeleccionar.AppendLine("           Prospectos, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Proposiciones.Documento  =   Renglones_Proposiciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Pro    =   Prospectos.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art		 =   Renglones_Proposiciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
            
            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuesto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpuesto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribución de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpuesto = lcPorcentajesImpuesto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuesto = lcPorcentajesImpuesto & ")"
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace("(,", "(")
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".", ",")
			

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProposiciones_Prospectos", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text38"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfProposiciones_Prospectos.ReportSource = loObjetoReporte

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
' MAT: 16/02/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
