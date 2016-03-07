'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing.Layout

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFPagoProveedor_SoloCheque_SID"
'-------------------------------------------------------------------------------------------'
Partial Class fFPagoProveedor_SoloCheque_SID
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
        loConsulta.AppendLine("SELECT  Pagos.Cod_Pro                AS Cod_Pro,") 
        loConsulta.AppendLine("        Proveedores.Nom_Pro          AS Nom_Pro,") 
        loConsulta.AppendLine("        Proveedores.Rif              AS Rif,") 
        loConsulta.AppendLine("        Proveedores.Nit              AS Nit,") 
        loConsulta.AppendLine("        Proveedores.Dir_Fis          AS Dir_Fis,") 
        loConsulta.AppendLine("        Proveedores.Telefonos        AS Telefonos,") 
        loConsulta.AppendLine("        Proveedores.Fax              AS Fax,") 
        loConsulta.AppendLine("        Pagos.Documento              AS Documento,") 
        loConsulta.AppendLine("        Pagos.Fec_Ini                AS Fec_Ini,") 
        loConsulta.AppendLine("        Pagos.Fec_Fin                AS Fec_Fin,") 
        loConsulta.AppendLine("        Pagos.Mon_Bru			    AS Mon_Bru_Enc,") 
        loConsulta.AppendLine("        (Pagos.Mon_Des * -1)		    AS Mon_Des,") 
        loConsulta.AppendLine("        Pagos.Mon_Net			    AS Mon_Net_Enc,") 
        loConsulta.AppendLine("        (Pagos.Mon_Ret * -1)		    AS Mon_Ret_Enc,") 
        loConsulta.AppendLine("        Pagos.Comentario			    AS Comentario,") 
        loConsulta.AppendLine("        Renglones_Pagos.Cod_Tip      AS Cod_Tip,") 
        loConsulta.AppendLine("        Renglones_Pagos.Doc_Ori      AS Doc_Ori,") 
        loConsulta.AppendLine("        Renglones_Pagos.Renglon      AS Renglon,") 
        loConsulta.AppendLine("        (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Bru ELSE (Renglones_Pagos.Mon_Bru * -1) END)  AS  Mon_Bru, ") 
        loConsulta.AppendLine("        (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Imp ELSE (Renglones_Pagos.Mon_Imp * -1) END)  AS  Mon_Imp, ") 
        loConsulta.AppendLine("        (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Abo ELSE (Renglones_Pagos.Mon_Abo * -1) END)  AS  Mon_Abo, ") 
        loConsulta.AppendLine("        (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Net ELSE (Renglones_Pagos.Mon_Net * -1) END)  AS  Mon_Net_Ren, ") 
        loConsulta.AppendLine("        CAST('' AS VARCHAR(MAX))     AS  Mon_Let ") 
        loConsulta.AppendLine("FROM    Pagos") 
        loConsulta.AppendLine("    JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.Documento") 
        loConsulta.AppendLine("    JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro") 
		loConsulta.AppendLine("WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal )         
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
' RJG: 30/09/13: Creación de la clase														'
'-------------------------------------------------------------------------------------------'
