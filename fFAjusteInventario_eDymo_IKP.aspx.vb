'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing.Layout

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFAjusteInventario_eDymo_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fFAjusteInventario_eDymo_IKP
    Inherits vis2formularios.frmReporte

#Region "Declaraciones"
        

    Private Const KN_ANCHO As Double = 40 '41.275  ' 1+5/8" = 1.625in
    Private Const KN_ALTO As Double = 31 '.75    ' 1+1/4" = 1.25in
    Private Const KC_INCLUYE_IVA As String = "Incluye I.V.A."

    Private loFuenteArial6 As XFont = New XFont("Arial", 6, XFontStyle.Regular)
    Private loFuenteArial8 As XFont = New XFont("Arial", 8, XFontStyle.Regular)
    Private loFuenteArial13 As XFont = New XFont("Arial", 13, XFontStyle.Regular)
    Private loFuenteTahoma8 As XFont = New XFont("Tahoma", 6, XFontStyle.Regular)
    Private loFuenteCode128 As XFont = New XFont("Code 128", 21, XFontStyle.Regular, New XPdfFontOptions(PdfFontEmbedding.Always))
    Private loFuenteCode128b As XFont = New XFont("Code 128", 20, XFontStyle.Regular, New XPdfFontOptions(PdfFontEmbedding.Always))

    Private lnFactor As Double = 72/25.4'Transforma milimetros en puntos (72 dpi)

    Private loCero As New XRect(-1, -1, 2, 2)
    Private loPosCodigoBarra As New XRect(-1, 1, KN_ANCHO*lnFactor, 4*lnFactor)
    Private loPosCodigo As New XRect(0, 0, KN_ANCHO*lnFactor, 8*lnFactor)

    Private loPosNombre As New XRect(1*lnFactor, 8*lnFactor, (KN_ANCHO-5.5)*lnFactor, 3*lnFactor)
    Private loPosClipNombre As New XRect(0*lnFactor, 8*lnFactor, (KN_ANCHO-4.5)*lnFactor, 3.5*lnFactor)

    Private loPosPrecio As New XRect(0, 15*lnFactor, (KN_ANCHO-2)*lnFactor, 2*lnFactor)
    Private loPosIvaInc As New XRect(0, 17*lnFactor, (KN_ANCHO-2)*lnFactor, 2*lnFactor)
    Private loPosFecha As New XRect(0, 21*lnFactor, (KN_ANCHO-2)*lnFactor, 2*lnFactor)
    Private loPosSerialBarra As New XRect(0, 22*lnFactor, KN_ANCHO*lnFactor, 4*lnFactor)
    Private loPosSerialBarraX2 As New XRect(0, 22*2*lnFactor, KN_ANCHO*lnFactor, 4*lnFactor)
    Private loPosSerial As New XRect(0, 24*lnFactor, KN_ANCHO*lnFactor, 4*lnFactor)

#End Region

#Region "Métodos"

    Private Function mGenerarCadenaCode128B(lcCadena As String ) As String

        If String.IsNullOrEmpty(lcCadena) Then Return ""

        Dim lnSum As Integer = 104 'CODE128/SET B = ASCII(0204)
        Dim lnCheckValue As Integer 
        Dim lcCheckChar As Char

        For n As Integer = 0 To lcCadena.Length - 1 
            Dim c As Char = lcCadena(n)
            lnSum += (Strings.Asc(c)-32) * (n+1)
        Next n
        
        lnCheckValue = (lnSum Mod 103) 
        If (lnCheckValue <=94) then
            lcCheckChar = Strings.Chr(lnCheckValue+32)
        Else
            lcCheckChar = Strings.Chr(lnCheckValue+100)
        End If

        Return Strings.Chr(0204) &  lcCadena & lcCheckChar & Strings.Chr(0206)

    End Function

    Private Sub mImprimirEtiqueta(loDocumentoPDF As PdfDocument, lcCodigo As String, lcNombre As String, _
                                  lcSerial As String, lcFecha As String, lcTelefono As String, lcPrecio As String)

        Dim loPaginaPDF As PdfPage
        Dim laGrafico As XGraphics
        Dim loEstado As XGraphicsState

        loPaginaPDF = loDocumentoPDF.AddPage()
        
        loPaginaPDF.Width = New XUnit(KN_ANCHO, XGraphicsUnit.Millimeter)
        loPaginaPDF.Height = New XUnit(KN_ALTO, XGraphicsUnit.Millimeter)

        laGrafico = XGraphics.FromPdfPage(loPaginaPDF)
        
        laGrafico.DrawString(mGenerarCadenaCode128B(lcSerial), loFuenteCode128, XBrushes.Black, loPosCodigoBarra, XStringFormats.BottomCenter)
        laGrafico.DrawString(lcSerial, loFuenteArial8, XBrushes.Black, loPosCodigo, XStringFormats.BottomCenter)

        loEstado = laGrafico.Save()
        laGrafico.IntersectClip(loPosClipNombre)
        laGrafico.DrawString(lcNombre, loFuenteTahoma8, XBrushes.Black, loPosNombre, XStringFormats.TopLeft)
        laGrafico.Restore(loEstado)

        laGrafico.DrawString(lcPrecio, loFuenteArial13, XBrushes.Black, loPosPrecio, XStringFormats.BottomCenter)
        laGrafico.DrawString(KC_INCLUYE_IVA, loFuenteArial6, XBrushes.Black, loPosIvaInc, XStringFormats.BottomCenter)
        laGrafico.DrawString(lcFecha, loFuenteArial6, XBrushes.Black, loPosFecha, XStringFormats.BottomCenter)
        
        'laGrafico.DrawString(lcSerial, loFuenteArial8, XBrushes.Black, loPosSerial, XStringFormats.BottomCenter)
        laGrafico.DrawString(lcTelefono, loFuenteArial6, XBrushes.Black, loPosSerial, XStringFormats.BottomCenter)

        loEstado = laGrafico.Save()
        laGrafico.RotateTransform(-90)
        laGrafico.TranslateTransform(-(KN_ALTO-4)*lnFactor, + 2*lnFactor)
        Dim loFormato As XStringFormat = (New XStringFormat())
        loFormato.Alignment = XStringAlignment.Near
        loFormato.LineAlignment = XLineAlignment.Center
        laGrafico.DrawString(lcCodigo, loFuenteArial6, XBrushes.Black, loCero, loFormato)
        laGrafico.Restore(loEstado)

        laGrafico.RotateTransform(-90)
        laGrafico.TranslateTransform(-(KN_ALTO-1)*lnFactor, +(KN_ANCHO-5)*lnFactor)
        laGrafico.DrawString(mGenerarCadenaCode128B(lcCodigo), loFuenteCode128b, XBrushes.Black, loCero, XStringFormats.TopLeft)
        

    End Sub

#End Region


    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

    '*****************************************************************************
    ' Busca la información
    '*****************************************************************************
        Dim loConsulta As New StringBuilder()
        Dim llIncluyeImpuesto As Boolean = cusAdministrativo.goArticulo.mPrecioIncluyeImpuesto("precio1")
        Dim lcTasaAdicional AS String = goServicios.mObtenerCampoFormatoSQL(goMoneda.pnTasaMonedaAdicional,goServicios.enuOpcionesRedondeo.KN_RedondeoUniforme, 10)


        loConsulta.AppendLine("CREATE TABLE #tmpSeriales(  Fec_Ini DATETIME,")
        loConsulta.AppendLine("                            Renglon INTEGER,")
        loConsulta.AppendLine("                            Telefonos VARCHAR(100) COLLATE DATABASE_DEFAULT,")
        loConsulta.AppendLine("                            Cod_Art VARCHAR(30) COLLATE DATABASE_DEFAULT,")
        loConsulta.AppendLine("                            Nom_Art VARCHAR(100) COLLATE DATABASE_DEFAULT,")
        loConsulta.AppendLine("                            Can_Art1 DECIMAL(28,10),")
        loConsulta.AppendLine("                            Ren_Ser INTEGER,")
        loConsulta.AppendLine("                            Serial VARCHAR(30) COLLATE DATABASE_DEFAULT,")
        loConsulta.AppendLine("                            Asignado DECIMAL(28,10) DEFAULT (0),")
        loConsulta.AppendLine("                            Sin_Asignar DECIMAL(28,10) DEFAULT (0));")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("DECLARE @lcTelefono VARCHAR(100);")
        loConsulta.AppendLine("SET @lcTelefono = (")
        loConsulta.AppendLine("    SELECT TOP 1 telefonos ")
        loConsulta.AppendLine("    FROM Sucursales")
        loConsulta.AppendLine("    WHERE cod_suc = " & goServicios.mObtenerCampoFormatoSQL(goSucursal.pcCodigo) &");")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("INSERT INTO #tmpSeriales(Renglon, Fec_Ini, Telefonos, Cod_Art, Nom_Art, Can_Art1, Ren_Ser, Serial)")
        loConsulta.AppendLine("SELECT   A.Renglon                      AS Renglon,")
        loConsulta.AppendLine("         A.Fec_Ini                      AS Fec_Ini,")
        loConsulta.AppendLine("         @lcTelefono                    AS Telefonos,")
        loConsulta.AppendLine("         A.Cod_Art                      AS Cod_Art,")
        loConsulta.AppendLine("         A.Nom_Art                      AS Nom_Art, ")
        loConsulta.AppendLine("         A.Can_Art1                     AS Can_Art1, ")
        loConsulta.AppendLine("         seriales.renglon               AS Ren_Ser,")
        loConsulta.AppendLine("         COALESCE(seriales.serial, '')  AS Serial")
        loConsulta.AppendLine("FROM    (   ")
        loConsulta.AppendLine("            SELECT   Renglones_Ajustes.documento    AS Documento,")
        loConsulta.AppendLine("                     Renglones_Ajustes.Renglon      AS Renglon,")
        loConsulta.AppendLine("                     Ajustes.Fec_Ini                AS Fec_Ini,")
        loConsulta.AppendLine("                     Renglones_Ajustes.Cod_Art      AS Cod_Art,")
        loConsulta.AppendLine("                     Renglones_Ajustes.Notas        AS Nom_Art, ")
        loConsulta.AppendLine("                     Renglones_Ajustes.Can_Art1     AS Can_Art1")
        loConsulta.AppendLine("            FROM     Ajustes ")
        loConsulta.AppendLine("             JOIN    Renglones_Ajustes ")
        loConsulta.AppendLine("                ON   Renglones_Ajustes.Documento = Ajustes.Documento")
        loConsulta.AppendLine("            WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
        loConsulta.AppendLine("        ) AS A ")
        loConsulta.AppendLine("LEFT JOIN seriales ON seriales.ren_ent = A.Renglon")
        loConsulta.AppendLine("    AND (   (tip_ent = 'Ajustes' AND doc_ent = A.Documento)")
        loConsulta.AppendLine("        OR  (tip_sal = 'Ajustes' AND doc_sal = A.Documento))")
        loConsulta.AppendLine("ORDER BY A.Renglon, seriales.renglon;")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("UPDATE  #tmpSeriales ")
        loConsulta.AppendLine("SET     Asignado = T.Asignado")
        loConsulta.AppendLine("FROM (  SELECT  #tmpSeriales.Cod_Art,")
        loConsulta.AppendLine("                #tmpSeriales.Renglon,")
        loConsulta.AppendLine("                COUNT(*) AS Asignado")
        loConsulta.AppendLine("        FROM    #tmpSeriales")
        loConsulta.AppendLine("        WHERE   Serial > ''")
        loConsulta.AppendLine("        GROUP BY Renglon, Cod_Art")
        loConsulta.AppendLine("        ) AS T")
        loConsulta.AppendLine("WHERE   #tmpSeriales.Cod_Art = T.Cod_Art")
        loConsulta.AppendLine("    AND #tmpSeriales.Renglon = T.Renglon")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("UPDATE  #tmpSeriales ")
        loConsulta.AppendLine("SET     Sin_Asignar = #tmpSeriales.Can_Art1 - Asignado")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("SELECT  #tmpSeriales.Renglon, ")
        loConsulta.AppendLine("        #tmpSeriales.Fec_Ini, ")
        loConsulta.AppendLine("        #tmpSeriales.Telefonos, ")
        loConsulta.AppendLine("        #tmpSeriales.Cod_Art, ")
        loConsulta.AppendLine("        #tmpSeriales.Nom_Art, ")
        loConsulta.AppendLine("        (CASE WHEN (articulos.Pre_nac=1)")
        loConsulta.AppendLine("            THEN Articulos.Precio1 ")
        loConsulta.AppendLine("            ELSE ROUND(Articulos.Precio1*" & lcTasaAdicional & ", 2)")
        loConsulta.AppendLine("        END ) AS Precio1,  ")
        loConsulta.AppendLine("        #tmpSeriales.Serial,")
        loConsulta.AppendLine("        (  impuestos.por_imp1 + impuestos.por_imp2 + impuestos.por_imp3 ")
        loConsulta.AppendLine("         + impuestos.por_imp4 + impuestos.por_imp5 + impuestos.por_imp6 ")
        loConsulta.AppendLine("         + impuestos.por_imp7 + impuestos.por_imp8 + impuestos.por_imp9 ")
        loConsulta.AppendLine("         + impuestos.por_imp10) AS Por_Imp")
        loConsulta.AppendLine("FROM  #tmpSeriales")
        loConsulta.AppendLine("    JOIN articulos ON articulos.cod_art = #tmpSeriales.cod_art")
        loConsulta.AppendLine("    JOIN impuestos ON articulos.cod_imp = impuestos.cod_imp")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("SELECT   Cod_Art, Asignado, Sin_Asignar ")
        loConsulta.AppendLine("FROM     #tmpSeriales")
        loConsulta.AppendLine("GROUP BY Cod_Art, Asignado, Sin_Asignar")
        loConsulta.AppendLine("ORDER BY MIN(Renglon), Cod_Art;")
        loConsulta.AppendLine("")

        'Me.mEscribirConsulta(loConsulta.ToString())

        Dim laDatosReporte As DataSet = (New goDatos().mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes"))
        Dim laSeriales As DataTable = laDatosReporte.Tables(0)
        Dim laArticulos As DataTable = laDatosReporte.Tables(1)

        
    '*****************************************************************************
    ' Genera el PDF
    '*****************************************************************************
        Dim loDocumentoPDF As New PdfDocument()
        
        loDocumentoPDF.Info.Title = "Formato de Impresión de Etiquetas"
        
        Dim lcCodigo As String
        Dim lcNombre As String
        Dim lcSerial As String
        Dim lcFecha As String
        Dim lcPrecio As String
        Dim lcTelefono As String  = ""
        Dim lnSinAsignar AS Integer

        Dim lcFechaDoc As String = ""

        If (laSeriales.Rows.Count>0) Then
            lcTelefono = CStr(laSeriales.Rows(0).Item("Telefonos")).Trim()
            If String.IsNullOrEmpty(lcTelefono) Then 
                lcTelefono = goEmpresa.pcTelefonoEmpresa.Trim()
            End If
            lcFechaDoc = CDate(laSeriales.Rows(0).Item("Fec_Ini")).ToString("yy-mm-yyyy")
        End If

        For Each loArticulo As datarow In laArticulos.Rows
            lcCodigo = CStr(loArticulo("Cod_Art")).Trim()
            lnSinAsignar = CInt(loArticulo("Sin_Asignar"))
            
            Dim laAsignados() As DataRow = laSeriales.Select("serial>'' AND cod_art=" & goServicios.mObtenerCampoFormatoSQL(lcCodigo))

            For Each loAsignado As DataRow In laAsignados
                lcNombre = CStr(loAsignado("Nom_Art")).Trim()
                lcSerial = CStr(loAsignado("Serial")).Trim()
                lcFecha = "Fecha: " & CDate(loAsignado("Fec_Ini")).ToString("dd/MM/yy")

                Dim lnPrecio As Decimal = CDec(loAsignado("Precio1"))
                If Not llIncluyeImpuesto Then
                    lnPrecio = goServicios.mRedondearValor((1+CDec(loAsignado("Por_Imp"))/100D)*lnPrecio, 2)
                End If
                lcPrecio = "Bs. " & lnPrecio.ToString("0.00")

                Me.mImprimirEtiqueta(loDocumentoPDF, lcCodigo, lcNombre, lcSerial, lcFecha, lcTelefono, lcPrecio)

            Next loAsignado

            If lnSinAsignar>0 Then

                Dim loSinAsignar As DataRow = laSeriales.Select("cod_art=" & goServicios.mObtenerCampoFormatoSQL(lcCodigo))(0)
                For n As Integer = 1 To lnSinAsignar
                    lcNombre = CStr(loSinAsignar("Nom_Art")).Trim()
                    lcSerial = ""
                    lcFecha = "Fecha: " & CDate(loSinAsignar("Fec_Ini")).ToString("dd/MM/yy")

                    Dim lnPrecio As Decimal = CDec(loSinAsignar("Precio1"))
                    If Not llIncluyeImpuesto Then
                        lnPrecio = goServicios.mRedondearValor((1+CDec(loSinAsignar("Por_Imp"))/100D)*lnPrecio, 2)
                    End If
                    lcPrecio = "Bs. " & lnPrecio.ToString("0.00")

                    Me.mImprimirEtiqueta(loDocumentoPDF, lcCodigo, lcNombre, lcSerial, lcFecha, lcTelefono, lcPrecio)

                Next n

            End If 
        Next loArticulo 


        Dim lcArchivo As String =  "fFAjusteInventario_eDymo_IKP_" & lcFechaDoc 
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

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 02/08/13: Creación de la clase														'
'-------------------------------------------------------------------------------------------'
' RJG: 04/08/13: Continuación de la programación.                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 05/08/13: Corrección en el cálculo del caracter de checkeo de Cobe128B.          	'
'-------------------------------------------------------------------------------------------'
' RJG: 08/08/13: Ajuste en el cálculo del caracter de checkeo de Code128B.          	    '
'-------------------------------------------------------------------------------------------'
' RJG: 09/08/13: Ajuste para mostrar precio con IVA.                                  	    '
'-------------------------------------------------------------------------------------------'
