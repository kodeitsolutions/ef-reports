'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fGuias_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class fGuias_Clientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Guias.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Guias.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Guias.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Guias.Rif = '') THEN Clientes.Rif ELSE Guias.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Guias.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Guias.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Guias.Telefonos = '') THEN Clientes.Telefonos ELSE Guias.Telefonos END) END) AS  Telefonos, ")
            
            
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Guias.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Guias.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Guias.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Guias.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Guias.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Guias.Documento, ")
            loComandoSeleccionar.AppendLine("           Guias.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Guias.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Guias.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Guias.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Guias.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Guias.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Guias.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Por_des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Guias, ")
            loComandoSeleccionar.AppendLine("           Renglones_Guias, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Guias.Documento  =   Renglones_Guias.Documento AND ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Guias.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Guias.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fGuias_Clientes", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text24"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfGuias_Clientes.ReportSource = loObjetoReporte

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
' JJD: 13/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 24/12/09: Se le incluyo los montos de descuentos y recargos.
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 21/04/10: Se ajusto para que tome el cliente generico
'-------------------------------------------------------------------------------------------'