Imports System.Data
Partial Class fLibres_Inventarios

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Inventarios.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Imp          As  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Status, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Comentario, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Cod_Mon          As  Moneda, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Tasa, ")            
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_uni1 AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Mon_Net      As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Por_Imp1     As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Imp      As  Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Mon_Imp1     As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Inventarios ")
            loComandoSeleccionar.AppendLine("           JOIN Renglones_LInventarios ON Renglones_LInventarios.Documento = Libres_Inventarios.Documento ")
			loComandoSeleccionar.AppendLine("           JOIN Articulos ON Articulos.Cod_Art = Renglones_LInventarios.Cod_Art ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Libres_Inventarios.Documento, Renglones_LInventarios.Renglon ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
     
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
                    'Verifica si el impuesto es igual a Cero
 						if CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
							lcPorcentajesImpuesto = lcPorcentajesImpuesto & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText),goServicios.enuOpcionesRedondeo.KN_RedondeoSuperior,2) & "%"
						End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuesto = lcPorcentajesImpuesto & ")"
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace("(,","(")
            
            if lcPorcentajesImpuesto = "()" Then
					lcPorcentajesImpuesto = " "
			End If


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLibres_Inventarios", laDatosReporte)
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".",",")
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLibres_Inventarios.ReportSource = loObjetoReporte

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
' JJD: 17/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 05/12/09: Ajuste para que se trajera la unidad desde Articulos y no desde el Renglon
'                del documento
'-------------------------------------------------------------------------------------------'
' CMS: 01/06/10: Se coloco la validación de registros 0 y el metodo de carga de imagen		'
'-------------------------------------------------------------------------------------------'
' MAT: 10/11/10: Mantenimiento del Reporte (Parte del Impuesto)
'-------------------------------------------------------------------------------------------'
