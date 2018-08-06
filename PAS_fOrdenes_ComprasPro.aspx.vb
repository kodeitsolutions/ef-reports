Imports System.Data
Partial Class PAS_fOrdenes_ComprasPro
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenCompra As String = cusAplicacion.goFormatos.pcCondicionPrincipal
        lcOrdenCompra = lcOrdenCompra.Substring(lcOrdenCompra.IndexOf("'"), (lcOrdenCompra.LastIndexOf("'") - lcOrdenCompra.IndexOf("'")) + 1)

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Control, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT Factory_Global.dbo.Usuarios.Nom_Usu ")
            loComandoSeleccionar.AppendLine("       FROM Factory_Global.dbo.Usuarios ")
            loComandoSeleccionar.AppendLine("       WHERE Factory_Global.dbo.Usuarios.Cod_Usu COLLATE SQL_Latin1_General_CP1_CI_AS = Ordenes_Compras.Usu_Cre COLLATE SQL_Latin1_General_CP1_CI_AS),'') AS Usuario_Crea,")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT Nom_Usu ")
            loComandoSeleccionar.AppendLine("                 FROM Factory_Global.dbo.Usuarios ")
            loComandoSeleccionar.AppendLine("                 WHERE Cod_Usu COLLATE DATABASE_DEFAULT = (SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("                                                           FROM Auditorias")
            loComandoSeleccionar.AppendLine("                                                           WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("                                                               AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("                                                               AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("                                                          ORDER BY Auditorias.Registro DESC)) COLLATE DATABASE_DEFAULT")
            loComandoSeleccionar.AppendLine("      ,'')	                            AS Usuario_Confirma,")
            loComandoSeleccionar.AppendLine("       Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art               AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Generico              AS Generico,")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Notas        AS Notas,")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Registro        As Fec_Cre, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Aut1        As Fec_Aut ")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON  Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   JOIN Formas_Pagos ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios5 As New cusDatos.goDatos

            Dim laDatosReporte5 As DataSet = loServicios5.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte5.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte5.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte5.Tables(0).Rows.Count - 1 Then
                        If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
                            lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                        End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte5.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte5.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PAS_fOrdenes_ComprasPro", laDatosReporte5)
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvPAS_fOrdenes_ComprasPro.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera datos de genericos de la Cotizacion cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 02/09/11: Adición de Comentario en Renglones
'-------------------------------------------------------------------------------------------'
