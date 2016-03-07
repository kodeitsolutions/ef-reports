'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFaltantes_Renuncia"
'-------------------------------------------------------------------------------------------'
Partial Class fFaltantes_Renuncia

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Faltantes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Documento, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,35)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Can_Pen1 AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Mon_Imp1         As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes.Por_Des AS Por_Des_Renglon, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Por_Des1,")
            loComandoSeleccionar.AppendLine("           Faltantes.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Faltantes.Mon_Des1,")
            loComandoSeleccionar.AppendLine("           Faltantes.Mon_Rec1 ")
            loComandoSeleccionar.AppendLine(" FROM      Faltantes, ")
            loComandoSeleccionar.AppendLine("           Renglones_Faltantes, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Faltantes.Documento =   Renglones_Faltantes.Documento AND ")
            loComandoSeleccionar.AppendLine("           Faltantes.Cod_Cli   =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Faltantes.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Faltantes.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   Renglones_Faltantes.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFaltantes_Renuncia", laDatosReporte)
            lcPorcentajesImpueto= lcPorcentajesImpueto.Replace(".", ",")            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFaltantes_Renuncia.ReportSource = loObjetoReporte

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
' CMS: 24/09/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 18/05/10: Se cambio el origen de datos de la columna Can_Art1: (Renglones_Faltantes.Can_Art1 - Renglones_Faltantes.Can_Pen1) 
'					a solo Renglones_Faltantes.Can_Pen1
'-------------------------------------------------------------------------------------------'
