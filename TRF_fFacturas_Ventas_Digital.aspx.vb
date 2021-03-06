﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_fFacturas_Ventas_Digital"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_fFacturas_Ventas_Digital
    Inherits vis2formularios.frmReporte 'Automatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT   Facturas.Cod_Cli                    AS Cod_Cli,")
            loConsulta.AppendLine("         Clientes.Nom_Cli					AS Nom_Cli, ")
            loConsulta.AppendLine("         Clientes.Rif						AS Rif, ")
            loConsulta.AppendLine("         Clientes.Dir_Fis	                AS Dir_Fis, ")
            loConsulta.AppendLine("         Clientes.Telefonos					AS Telefonos, ")
            loConsulta.AppendLine("         Facturas.Documento                  AS Documento, ")
            loConsulta.AppendLine("         Facturas.Control                    AS Control, ")
            loConsulta.AppendLine("         Facturas.Fec_Ini                    AS Fec_Ini, ")
            loConsulta.AppendLine("         Facturas.Fec_Fin                    AS Fec_Fin, ")
            loConsulta.AppendLine("         Facturas.Mon_Bru                    AS Mon_Bru, ")
            loConsulta.AppendLine("         Facturas.Mon_Exe                    AS Mon_Exe, ")
            loConsulta.AppendLine("         Facturas.Mon_Imp1                   AS Mon_Imp1, ")
            loConsulta.AppendLine("         Facturas.Por_Imp1                   AS Por_Imp1, ")
            loConsulta.AppendLine("         Facturas.Dis_Imp                    AS Dis_Imp, ")
            loConsulta.AppendLine("         Facturas.Mon_Net                    AS Mon_Net, ")
            loConsulta.AppendLine("         Facturas.Cod_For                    AS Cod_For, ")
            loConsulta.AppendLine("         Formas_Pagos.Nom_For AS Nom_For, ")
            loConsulta.AppendLine("         Facturas.Comentario                 AS Comentario,")
            loConsulta.AppendLine("         Renglones_Facturas.Cod_Art          AS Cod_Art, ")
            loConsulta.AppendLine("         CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("             THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("		    ELSE Renglones_Facturas.Notas END   AS Nom_Art, ")
            loConsulta.AppendLine("         Renglones_Facturas.Can_Art1			AS Can_Art1, ")
            loConsulta.AppendLine("         Renglones_Facturas.Cod_Uni			AS Cod_Uni, ")
            loConsulta.AppendLine("         Renglones_Facturas.Precio1			AS Precio1,")
            loConsulta.AppendLine("         Renglones_Facturas.Mon_Net          AS Neto, ")
            loConsulta.AppendLine("         Renglones_Facturas.Por_Imp1         AS Por_Imp, ")
            loConsulta.AppendLine("         Renglones_Facturas.Cod_Imp          AS Cod_Imp, ")
            loConsulta.AppendLine("         Renglones_Facturas.Mon_Imp1         AS Impuesto ")
            loConsulta.AppendLine("FROM Facturas ")
            loConsulta.AppendLine("    JOIN Renglones_Facturas ON Facturas.Documento = Renglones_Facturas.Documento")
            loConsulta.AppendLine("    JOIN Clientes ON  Facturas.Cod_Cli = Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN Formas_Pagos ON  Facturas.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN Articulos ON  Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuesto As String = ""
            Dim loImpuestos As New System.Xml.XmlDocument()

            'lcPorcentajesImpuesto = "("

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
                        If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
                            lcPorcentajesImpuesto = lcPorcentajesImpuesto & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                        End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("TRF_fFacturas_Ventas_Digital", laDatosReporte)
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".", ",")
            lcPorcentajesImpuesto = "I.V.A.: " & lcPorcentajesImpuesto
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvTRF_fFacturas_Ventas_Digital.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' GMO: 16/08/08: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' JJD: 08/11/08: Ajustes al select.                                                         '
'-------------------------------------------------------------------------------------------'
' RJG: 01/09/09: Agregado código para mostrar unidad segundaria.							'
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen. '
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera los datos genericos de la Factura cuando aplique.'
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero.    '
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: Se a justo la logica para determinar el nombre del cliente                 '
'		(Clientes.Generico = 0 ) a (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '')  '
'-------------------------------------------------------------------------------------------'
' MAT: 23/02/11: Se programo la distribución de impuestos para mostrarlo en el formato.		'
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.                                            '
'-------------------------------------------------------------------------------------------'
' RJG:  06/02/14 : Ajuste de la formato (comentario, indentación...) de código y SQL.       '
'-------------------------------------------------------------------------------------------'
