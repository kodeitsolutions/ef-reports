'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing.Layout

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFOrdenPagoProveedor_SoloCheque_SID"
'-------------------------------------------------------------------------------------------'
Partial Class fFOrdenPagoProveedor_SoloCheque_SID
    Inherits vis2formularios.frmReporte

#Region "Declaraciones"
        

#End Region

#Region "Métodos"

#End Region

#Region "Eventos"

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

    '*****************************************************************************
    ' Busca la información
    '*****************************************************************************
        Dim loConsulta As New StringBuilder()
        Dim llIncluyeImpuesto As Boolean = cusAdministrativo.goArticulo.mPrecioIncluyeImpuesto("precio1")
        Dim lcTasaAdicional AS String = goServicios.mObtenerCampoFormatoSQL(goMoneda.pnTasaMonedaAdicional,goServicios.enuOpcionesRedondeo.KN_RedondeoUniforme, 10)
        
        loConsulta.AppendLine("") 
            loConsulta.AppendLine(" SELECT	    Ordenes_Pagos.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Pagos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Pagos.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Pagos.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax                               AS Fax, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Nom_Pro                         AS  Nombre_Generico, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Rif                             AS  Rif_Genenerico, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Nit                             AS  Nit_Generico, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Dir_Fis                         AS  Dir_Fis_Generico, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Telefonos                       AS  Telefonos_Generico, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Documento                       AS  Documento, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Fec_Ini                         AS  Fec_Ini, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Fec_Fin                         AS  Fec_Fin, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Mon_Bru                         AS  Mon_Bru_Enc, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Mon_Imp                         AS  Mon_Imp1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Mon_Net                         AS  Mon_Net_Enc, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Mon_Ret                         AS  Mon_Ret_Enc, ")
            loConsulta.AppendLine("           Ordenes_Pagos.Motivo                          AS  Motivo, ")
            loConsulta.AppendLine("           Renglones_oPagos.Cod_Con, ")
            loConsulta.AppendLine("           Conceptos.Nom_Con + Substring(Renglones_oPagos.Comentario,1,250)    As  Nom_Con, ")
            loConsulta.AppendLine("           Renglones_oPagos.Renglon, ")
            loConsulta.AppendLine("           Renglones_oPagos.Mon_Deb                      AS  Mon_Deb, ")
            loConsulta.AppendLine("           Renglones_oPagos.Mon_Hab                      AS  Mon_Hab, ")
            loConsulta.AppendLine("            CASE ")
            loConsulta.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Net * -1 ")
            loConsulta.AppendLine("            	ELSE ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Net ")
            loConsulta.AppendLine("            END                                          AS  Mon_Net_Ren, ")
            loConsulta.AppendLine("            CASE ")
            loConsulta.AppendLine("            	WHEN Renglones_oPagos.Mon_Deb = 0 THEN ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Imp1 * -1 ")
            loConsulta.AppendLine("            	ELSE ")
            loConsulta.AppendLine("            		Renglones_oPagos.Mon_Imp1 ")
            loConsulta.AppendLine("            END                                          AS  Mon_Imp_Ren, ")
            loConsulta.AppendLine("           Renglones_oPagos.Cod_Imp                      AS  Cod_Imp_Ren, ")
            loConsulta.AppendLine("           Renglones_oPagos.Comentario                   AS  Comentario_Ren, ")
            loConsulta.AppendLine("           Renglones_oPagos.Mon_Imp1                     AS  Mon_Imp_Ren, ")
            loConsulta.AppendLine("           CAST('' AS VARCHAR(MAX))                      AS  Mon_Let ")
            loConsulta.AppendLine(" FROM      Ordenes_Pagos, ")
            loConsulta.AppendLine("           Renglones_oPagos, ")
            loConsulta.AppendLine("           Proveedores, ")
            loConsulta.AppendLine("           Conceptos ")
            loConsulta.AppendLine(" WHERE     Ordenes_Pagos.Documento =   Renglones_oPagos.Documento")
            loConsulta.AppendLine("     AND   Ordenes_Pagos.Cod_Pro   =   Proveedores.Cod_Pro ")
            loConsulta.AppendLine("     AND   Conceptos.Cod_Con       =   Renglones_oPagos.Cod_Con ")
		    loConsulta.AppendLine("     AND   " & cusAplicacion.goFormatos.pcCondicionPrincipal )         
            loConsulta.AppendLine("") 
        
        'Me.mEscribirConsulta(loConsulta.ToString)

        Dim laDatosReporte As DataTable = (New goDatos().mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")).Tables(0)

        If laDatosReporte.Rows.Count = 0 Then 

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                "El documento seleccionado no existe o no es válido.", _
                vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                "auto", _
                "auto")

            Return

        End If

        Dim loCheque As DataRow = laDatosReporte.Rows(0)
        
    '*****************************************************************************
    ' Prepara los datos para el PDF
    '*****************************************************************************
        Dim loDocumentoPDF As New PdfDocument()
 
        Const KN_ANCHO As Double = 172 'milímetros
        Const KN_ALTO As Double = 78 'milímetros
        Const KN_FACTOR As Double = 72/25.4 'Transforma milimetros en puntos (72 dpi)

        Dim loFuenteArial_11 As XFont = New XFont("Arial", 11, XFontStyle.Bold)
        Dim loFuenteArialNarrow_09 As XFont = New XFont("Arial Narrow", 9, XFontStyle.Regular)
        Dim loFuenteArial_08 As XFont = New XFont("Arial", 9, XFontStyle.Regular)
       
        Dim lcTituloFormato     As String = "Formato de Cheque Banesco (SID - Solo Cheque)"
        Dim lcNombreProveedor   As String = CStr(loCheque("Nom_Pro")).Trim()
        Dim lnMontoCheque       As Decimal = CDec(loCheque("Mon_Net_Enc"))
        Dim lcMontoCheque       As String = "***" & goServicios.mObtenerFormatoCadena(lnMontoCheque,goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)
        Dim lcMontoLetras       As String = goServicios.mConvertirMontoLetras(lnMontoCheque)
        Dim lcFechaEmision      As String = CDate(loCheque("Fec_Ini")).ToString("dd-MM-yyyy")
        Dim lcNoEndosable       As String = "NO ENDOSABLE"

    '*****************************************************************************
    ' Genera el PDF
    '*****************************************************************************
        Dim loPaginaPDF As PdfPage
        Dim laGrafico As XGraphics

        Dim loPosMonto As New XRect(130*KN_FACTOR, 9*KN_FACTOR, 30*KN_FACTOR, 0*KN_FACTOR)
        Dim loPosNombre As New XRect(25*KN_FACTOR, 22*KN_FACTOR, 140*KN_FACTOR, 0*KN_FACTOR)
        Dim loPosLetras As New XRect(25*KN_FACTOR, 27*KN_FACTOR, 140*KN_FACTOR, 0*KN_FACTOR)
        Dim loPosFecha As New XRect(35*KN_FACTOR, 39*KN_FACTOR, 20*KN_FACTOR, 0*KN_FACTOR)
        Dim loPosNoEndo As New XRect(85*KN_FACTOR, 63*KN_FACTOR, 40*KN_FACTOR, 0*KN_FACTOR)

        loDocumentoPDF.Info.Title = lcTituloFormato

        loPaginaPDF = loDocumentoPDF.AddPage()
        loPaginaPDF.Width = New XUnit(KN_ANCHO, XGraphicsUnit.Millimeter)
        loPaginaPDF.Height = New XUnit(KN_ALTO, XGraphicsUnit.Millimeter)

        laGrafico = XGraphics.FromPdfPage(loPaginaPDF)
        laGrafico.DrawString(lcMontoCheque, loFuenteArial_11, XBrushes.Black, loPosMonto, XStringFormats.Default)
        laGrafico.DrawString(lcNombreProveedor, loFuenteArial_11, XBrushes.Black, loPosNombre, XStringFormats.Default)
        laGrafico.DrawString(lcMontoLetras, loFuenteArialNarrow_09, XBrushes.Black, loPosLetras, XStringFormats.Default)
        laGrafico.DrawString(lcFechaEmision, loFuenteArial_08, XBrushes.Black, loPosFecha, XStringFormats.Default)
        laGrafico.DrawString(lcNoEndosable, loFuenteArial_11, XBrushes.Black, loPosNoEndo, XStringFormats.Default)

    '*****************************************************************************
    ' Descarga el archivo PDF
    '*****************************************************************************

        Dim lcArchivo As String =  "fFNotasRecepcion_eDymo_IKP_" & lcFechaEmision 
        Dim lcRuta As String = Me.Server.MapPath("~/Administrativo/Temporales/" & lcArchivo & "_" & Guid.NewGuid().ToString("N").ToUpper().Substring(0, 10) & ".pdf")
        Try 
            If (My.Computer.FileSystem.FileExists(lcArchivo)) Then 
                My.Computer.FileSystem.DeleteFile(lcRuta)
            End If
            loDocumentoPDF.Save(lcRuta)

            Me.Response.Clear()
            Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcArchivo)
            Me.Response.ContentType = "application/pdf"
            Me.Response.WriteFile(lcRuta)
            Me.Response.Flush()
            Me.Response.End()

        Catch ex As Exception
            
        End Try

    End Sub

#End Region

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/10/13: Creación de la clase														'
'-------------------------------------------------------------------------------------------'
