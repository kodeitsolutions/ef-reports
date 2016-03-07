'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_TicketIpos"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_TicketIpos
    Inherits vis2formularios.frmReporte

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        loConsulta.AppendLine("")
        loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
        loConsulta.AppendLine("SET @lcDocumento = (")
        loConsulta.AppendLine("    SELECT TOP 1 Documento")
        loConsulta.AppendLine("    FROM Facturas ")
        loConsulta.AppendLine("    WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal & ");")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("SELECT      Facturas.Cod_Cli                            AS Cod_Cli, ")
        loConsulta.AppendLine("            (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) AS Nom_Cli, ")
        loConsulta.AppendLine("            (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) AS  Rif, ")
        loConsulta.AppendLine("            Clientes.Nit                                AS Nit, ")
        loConsulta.AppendLine("            (CASE WHEN (Facturas.Dir_Fis = '') THEN Clientes.Dir_Fis ELSE Facturas.Dir_Fis END) AS Dir_Fis, ")
        loConsulta.AppendLine("            (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) AS  Telefonos, ")
        loConsulta.AppendLine("            Clientes.Fax                                AS Fax, ")
        loConsulta.AppendLine("            Clientes.Generico                           AS Generico, ")
        loConsulta.AppendLine("            Facturas.Documento                          AS Documento, ")
        loConsulta.AppendLine("            Facturas.Fec_Ini                            AS Fec_Ini, ")
        loConsulta.AppendLine("            Facturas.Fec_Fin                            AS Fec_Fin, ")
        loConsulta.AppendLine("            Facturas.Cod_Mon                            AS Cod_Mon, ")
        loConsulta.AppendLine("            Facturas.Tasa                               AS Tasa, ")
        loConsulta.AppendLine("            Facturas.Mon_Bru                            AS Mon_Bru, ")
        loConsulta.AppendLine("            Facturas.Mon_Imp1                           AS Mon_Imp1, ")
        loConsulta.AppendLine("            Facturas.Por_Imp1                           AS Por_Imp1, ")
        loConsulta.AppendLine("            Facturas.Mon_Net                            AS Mon_Net, ")
        loConsulta.AppendLine("            Facturas.Dis_Imp                            AS Dis_Imp, ")
        loConsulta.AppendLine("            Facturas.Por_Des1                           AS Por_Des, ")
        loConsulta.AppendLine("            Facturas.Mon_Des1                           AS Mon_Des, ")
        loConsulta.AppendLine("            Facturas.Por_Rec1                           AS Por_Rec, ")
        loConsulta.AppendLine("            Facturas.Mon_Rec1                           AS Mon_Rec, ")
        loConsulta.AppendLine("            Facturas.Cod_For                            AS Cod_For, ")
        loConsulta.AppendLine("            Formas_Pagos.Nom_For                        AS Nom_For, ")
        loConsulta.AppendLine("            Facturas.Cod_Ven                            AS Cod_Ven, ")
        loConsulta.AppendLine("            Facturas.Comentario                         AS Comentario, ")
        loConsulta.AppendLine("            Vendedores.Nom_Ven                          AS Nom_Ven, ")
        loConsulta.AppendLine("            Renglones_Facturas.Renglon                  AS Renglon, ")
        loConsulta.AppendLine("            Renglones_Facturas.Cod_Art                  AS Cod_Art, ")
        loConsulta.AppendLine("            Renglones_Facturas.Notas                    AS Nom_Art,  ")
        loConsulta.AppendLine("            Renglones_Facturas.Can_Art1                 AS Can_Art1, ")
        loConsulta.AppendLine("            Renglones_Facturas.Cod_Uni                  AS Cod_Uni, ")
        loConsulta.AppendLine("            Renglones_Facturas.Precio1                  AS Precio1,")
        loConsulta.AppendLine("            Renglones_Facturas.Mon_Net                  AS Neto, ")
        loConsulta.AppendLine("            Renglones_Facturas.Por_Imp1                 AS Por_Imp, ")
        loConsulta.AppendLine("            Renglones_Facturas.Cod_Imp                  AS Cod_Imp, ")
        loConsulta.AppendLine("            Renglones_Facturas.Mon_Imp1                 AS Impuesto")
        loConsulta.AppendLine("FROM        Facturas ")
        loConsulta.AppendLine("    JOIN    Renglones_Facturas")
        loConsulta.AppendLine("        ON  Facturas.Documento  =   Renglones_Facturas.Documento")
        loConsulta.AppendLine("    JOIN    Clientes")
        loConsulta.AppendLine("        ON  Facturas.Cod_Cli    =   Clientes.Cod_Cli")
        loConsulta.AppendLine("    JOIN    Formas_Pagos")
        loConsulta.AppendLine("        ON  Facturas.Cod_For    =   Formas_Pagos.Cod_For")
        loConsulta.AppendLine("    JOIN    Vendedores ")
        loConsulta.AppendLine("        ON  Facturas.Cod_Ven    =   Vendedores.Cod_Ven")
        loConsulta.AppendLine("    JOIN    Articulos ")
        loConsulta.AppendLine("        ON  Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art")
        loConsulta.AppendLine("WHERE       Facturas.Documento  =  @lcDocumento")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("DECLARE @lxImpuestos AS XML;")
        loConsulta.AppendLine("SET @lxImpuestos = (SELECT TOP 1 Facturas.Dis_Imp FROM Facturas WHERE Facturas.Documento = @lcDocumento);")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("SELECT  T.C.value('./porcentaje[1]', 'DECIMAL(28,10)')  porcentaje,")
        loConsulta.AppendLine("        T.C.value('./base[1]', 'DECIMAL(28,10)')  base,")
        loConsulta.AppendLine("        T.C.value('./exento[1]', 'DECIMAL(28,10)')  exento,")
        loConsulta.AppendLine("        T.C.value('./monto[1]', 'DECIMAL(28,10)')  monto")
        loConsulta.AppendLine("FROM    Facturas ")
        loConsulta.AppendLine("CROSS APPLY @lxImpuestos.nodes('//impuestos/impuesto') T(C)")
        loConsulta.AppendLine("WHERE Facturas.Documento  =  @lcDocumento")
        loConsulta.AppendLine("ORDER BY porcentaje DESC")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")

        'Me.mEscribirTexto(loConsulta.ToString())

        Dim loServicios As New cusDatos.goDatos()

        Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

        '--------------------------------------------------'
        ' Carga la imagen del logo en cusReportes            '
        '--------------------------------------------------'
        'Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

        Me.mGenerarTicket(laDatosReporte.Tables(0), laDatosReporte.Tables(1))

    End Sub

    Private Sub mGenerarTicket(loRenglones As DataTable, loImpuestos As DataTable)

        'Datos de la factura:
        Dim loFactura As DataRow = loRenglones.Rows(0)
        Dim lcFactura As String = CStr(loFactura("Documento")).Trim()
        Dim lcMoneda As String = "$ " 'CStr(loFactura("Cod_Mon")).Trim()
        Dim lcCliente_Rif As String = CStr(loFactura("Rif")).Trim()
        Dim lcCliente_Nombre As String = CStr(loFactura("Nom_Cli")).Trim()
        Dim lcCliente_Direccion As String = CStr(loFactura("Dir_Fis")).Trim()
        Dim lcCliente_Telefono As String = CStr(loFactura("Telefonos")).Trim()

    '*****************************************************
    ' Definición del PDF (documento y valores)
    '*****************************************************
        Dim loDocumento As New PdfDocument()
        Dim loPagina As PdfPage = loDocumento.AddPage()
        Dim loGrafico As XGraphics = XGraphics.FromPdfPage(loPagina)

        'Define el ancho de la página
        loPagina.Width = XUnit.FromMillimeter(60)

        Dim lnAnchoPagina As Double = CDbl(loPagina.Width)
        Dim lnInternineado As Double = XUnit.FromMillimeter(2)
        Dim lnMargenV As Double = XUnit.FromMillimeter(5)
        Dim lnMargen As Double = XUnit.FromMillimeter(3)
        Dim lnMargenC As Double = XUnit.FromMillimeter(1)

        Dim loFuenteTitulo As XFont = new XFont("Courier New", 8, XFontStyle.Bold)
        Dim loFuente As XFont = new XFont("Courier New", 6, XFontStyle.Regular)
        
        Dim loLineaTitulo As XRect = New XRect(0, 0, lnAnchoPagina, loFuenteTitulo.Height)
        Dim loLinea As XRect = New XRect(lnMargen, 0, lnAnchoPagina-lnMargen*2, loFuente.Height)
        Dim loLineaDerecha As XRect = New XRect(loLinea.X, loLinea.Y, 0, loLinea.Height) 'Usada para alinear a la derecha

        Dim loTrazo As New XPen(XColor.FromArgb(0), 0.5) 'Para dibujar línea horizontal de separación

        Dim lcTexto As String 'Auxiliar para texto a imprimir
        
    '*****************************************************
    ' Encabezado
    '*****************************************************
        loLineaTitulo.Offset(0, lnMargenV)

        loGrafico.DrawString(goEmpresa.pcNombre, loFuenteTitulo, XBrushes.Black, loLineaTitulo, XStringFormats.Center)
        loLineaTitulo.Offset(0, loLineaTitulo.Height)
        
        loGrafico.DrawString(goEmpresa.pcDireccionEmpresa, loFuenteTitulo, XBrushes.Black, loLineaTitulo, XStringFormats.Center)
        loLineaTitulo.Offset(0, loLineaTitulo.Height)
        
        loGrafico.DrawString(goEmpresa.pcTelefonoEmpresa, loFuenteTitulo, XBrushes.Black, loLineaTitulo, XStringFormats.Center)
        loLineaTitulo.Offset(0, loLineaTitulo.Height)

        loLinea.Offset(0, loLineaTitulo.Location.Y)
        loGrafico.DrawString(goEmpresa.pcRifEmpresa, loFuente, XBrushes.Black, loLinea, XStringFormats.Center)
        loLinea.Offset(0, loLineaTitulo.Height)
        
        'Separador
        loLinea.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )

    '*****************************************************
    ' Número de Factura, Fecha y Hora
    '*****************************************************
        loLinea.Offset(0, lnMargenC)
        loGrafico.DrawString("FECHA: " & Date.Now().ToString("dd-MM-yyyy"), loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
 
        lcTexto = "HORA: " & Date.Now().ToString("HH:mm")
        loLineaDerecha.Y = loLinea.Y
        loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
        loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
        loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

        loLinea.Offset(0, loLinea.Height)
        loGrafico.DrawString("FACTURA #: ", loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
 
        lcTexto = lcFactura
        loLineaDerecha.Y = loLinea.Y
        loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
        loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
        loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

        loLinea.Offset(0, loLinea.Height)

        'Separador
        loLinea.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )

    '*****************************************************
    ' Datos del cliente
    '*****************************************************
        loLinea.Offset(0, lnMargenC)
        If Not String.IsNullOrEmpty(lcCliente_Rif) Then
            loGrafico.DrawString("CI\RIF: " & lcCliente_Rif, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
            loLinea.Offset(0, loLinea.Height)
        End If

        Dim lnLargoT As Double  
        If Not String.IsNullOrEmpty(lcCliente_Nombre) Then
            lcTexto = ""
            lcCliente_Nombre = "NOMBRE: " & lcCliente_Nombre
            lnLargoT = loGrafico.MeasureString(lcCliente_Nombre, loFuente).Width

            'Si no cabe en una línea: lo corta
            While ( (lnLargoT > lnAnchoPagina - lnMargen*2) AND lcCliente_Nombre.Length > 0)
                lcTexto = lcCliente_Nombre(lcCliente_Nombre.Length-1) & lcTexto
                lcCliente_Nombre = lcCliente_Nombre.Substring(0, lcCliente_Nombre.Length - 1)
                lnLargoT = loGrafico.MeasureString(lcCliente_Nombre, loFuente).Width
            End While

            loGrafico.DrawString(lcCliente_Nombre, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
            loLinea.Offset(0, loLinea.Height)

            'Si cortó el texto: imprimir una segunda línea (el resto se pierde)
            If (lcTexto.Length > 0)
                lcCliente_Nombre = Strings.Space(8) & lcTexto
                lnLargoT = loGrafico.MeasureString(lcCliente_Nombre, loFuente).Width

                While ( (lnLargoT > lnAnchoPagina - lnMargen*2) AND lcCliente_Nombre.Length > 0)
                    lcCliente_Nombre = lcCliente_Nombre.Substring(0, lcCliente_Nombre.Length - 1)
                    lnLargoT = loGrafico.MeasureString(lcCliente_Nombre, loFuente).Width
                End While

                loGrafico.DrawString(lcCliente_Nombre, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
                loLinea.Offset(0, loLinea.Height)

            End If

        End If

        If Not String.IsNullOrEmpty(lcCliente_Direccion) Then
            lcTexto = ""
            lcCliente_Direccion = "DIRECCCIÓN: " & lcCliente_Direccion
            lnLargoT = loGrafico.MeasureString(lcCliente_Direccion, loFuente).Width

            'Si no cabe en una línea: lo corta
            While ( (lnLargoT > lnAnchoPagina - lnMargen*2) AND lcCliente_Direccion.Length > 0)
                lcTexto = lcCliente_Direccion(lcCliente_Direccion.Length-1) & lcTexto
                lcCliente_Direccion = lcCliente_Direccion.Substring(0, lcCliente_Direccion.Length - 1)
                lnLargoT = loGrafico.MeasureString(lcCliente_Direccion, loFuente).Width
            End While

            loGrafico.DrawString(lcCliente_Direccion, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
            loLinea.Offset(0, loLinea.Height)

            'Si cortó el texto: imprimir una segunda línea (el resto se pierde)
            If (lcTexto.Length > 0)
                lcCliente_Direccion = Strings.Space(8) & lcTexto
                lnLargoT = loGrafico.MeasureString(lcCliente_Direccion, loFuente).Width

                While ( (lnLargoT > lnAnchoPagina - lnMargen*2) AND lcCliente_Direccion.Length > 0)
                    lcCliente_Direccion = lcCliente_Direccion.Substring(0, lcCliente_Direccion.Length - 1)
                    lnLargoT = loGrafico.MeasureString(lcCliente_Direccion, loFuente).Width
                End While

                loGrafico.DrawString(lcCliente_Direccion, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
                loLinea.Offset(0, loLinea.Height)

            End If

        End If

        If Not String.IsNullOrEmpty(lcCliente_Telefono) Then
            lcTexto = ""
            lcCliente_Telefono = "TELÉFONO: " & lcCliente_Telefono
            lnLargoT = loGrafico.MeasureString(lcCliente_Telefono, loFuente).Width

            'Si no cabe en una línea: lo corta (el resto se pierde)
            While ( (lnLargoT > lnAnchoPagina - lnMargen*2) AND lcCliente_Telefono.Length > 0)
                lcTexto = lcCliente_Telefono(lcCliente_Telefono.Length-1) & lcTexto
                lcCliente_Telefono = lcCliente_Telefono.Substring(0, lcCliente_Telefono.Length - 1)
                lnLargoT = loGrafico.MeasureString(lcCliente_Telefono, loFuente).Width
            End While

            loGrafico.DrawString(lcCliente_Telefono, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)
            loLinea.Offset(0, loLinea.Height)

        End If

        'Separador
        loLinea.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )

    '*****************************************************
    ' Artículos
    '*****************************************************
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad

        loLinea.Offset(0, lnMargenC)
        Dim lnDisponibleArticulo As Double = XUnit.FromMillimeter(41).Point
        loLineaDerecha.Y = loLinea.Y
        For Each loRenglon As DataRow In loRenglones.Rows
            Dim lnCantidadArt As Decimal = CDec(loRenglon("Can_Art1"))
             
            'La cantidad solo se indica si es diferente a 1
            If (lnCantidadArt <> 1D)  Then

                If (lnCantidadArt = Decimal.Floor(lnCantidadArt)) OrElse (lnDecimalesCantidad = 0) Then
                    lcTexto = goServicios.mObtenerFormatoCadena(CInt(lnCantidadArt), 0) & _
                              " x " & lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loRenglon("Precio1")), 2)
                Else
                    lcTexto = goServicios.mObtenerFormatoCadena(lnCantidadArt, lnDecimalesCantidad) & _
                              " x " & lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loRenglon("Precio1")), 2)
                End If
                    

                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

                loLinea.Offset(0, loLinea.Height)
                loLineaDerecha.Offset(0, loLineaDerecha.Height)

            End If


            Dim lnAnchoTexto As Double 
            
            lcTexto = CStr(loRenglon("Nom_Art")).Trim()
            lnAnchoTexto = loGrafico.MeasureString(lcTexto, loFuente).Width
            While(lnAnchoTexto > lnDisponibleArticulo)
                lcTexto = lcTexto.Substring(0, lcTexto.Length -1)
                lnAnchoTexto = loGrafico.MeasureString(lcTexto, loFuente).Width
            End While
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

            lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loRenglon("Neto")), 2)
            loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
            loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

            loLinea.Offset(0, loLinea.Height)
            loLineaDerecha.Offset(0, loLineaDerecha.Height)

        Next

        'Separador
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)

    '*****************************************************
    ' Sub total, Descuentos y Recargos
    '*****************************************************
        loGrafico.DrawString("SUBTTL: ", loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

        lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Bru")), 2)
        loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
        loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
        loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

        loLinea.Offset(0, loLinea.Height)
        loLineaDerecha.Offset(0, loLineaDerecha.Height)

        'Descuento
        If (CDec(loFactura("Mon_Des"))>0D) Then
            lcTexto = "Descuento (" & goServicios.mObtenerFormatoCadena(CDec(loFactura("Por_Des")), 2) & "%):"
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

            lcTexto = lcMoneda & "-" & goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Des")), 2)
            loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
            loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

            loLinea.Offset(0, loLinea.Height)
            loLineaDerecha.Offset(0, loLineaDerecha.Height)

        End If

        'Recargo
        If (CDec(loFactura("Mon_Rec"))>0D) Then
            lcTexto = "Recargo (" & goServicios.mObtenerFormatoCadena(CDec(loFactura("Por_Rec")), 2) & "%):"
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

            lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Rec")), 2)
            loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
            loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
            loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

            loLinea.Offset(0, loLinea.Height)
            loLineaDerecha.Offset(0, loLineaDerecha.Height)

        End If

        'Separador
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)

    '*****************************************************
    ' Detalle de Impuestos
    '*****************************************************
        For Each loImpuesto As DataRow In loImpuestos.Rows
            Dim lnPorcentaje As Decimal = CDec(loImpuesto("Porcentaje"))
            Dim lnBase As Decimal = CDec(loImpuesto("Base")) + CDec(loImpuesto("Exento")) 'Solo uno de los dos campos es >0
            Dim lnImpuesto As Decimal = CDec(loImpuesto("Monto"))

            If (lnPorcentaje > 0D) Then

                lcTexto = "BI.G.(" & goServicios.mObtenerFormatoCadena(lnPorcentaje, 2) & "%):"
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

                lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(lnBase, 2)
                loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
                loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)
                
                loLinea.Offset(0, loLinea.Height)
                loLineaDerecha.Offset(0, loLineaDerecha.Height)

                lcTexto = "IVA.G.(" & goServicios.mObtenerFormatoCadena(lnPorcentaje, 2) & "%):"
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

                lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(lnImpuesto, 2)
                loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
                loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

            Else

                lcTexto = "EXENTO:"
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLinea, XStringFormats.TopLeft)

                lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(lnBase, 2)
                loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuente).Width
                loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
                loGrafico.DrawString(lcTexto, loFuente, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

            End If

            loLinea.Offset(0, loLinea.Height)
            loLineaDerecha.Offset(0, loLineaDerecha.Height)

        Next loImpuesto

        'Separador
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)



    '*****************************************************
    ' Total
    '*****************************************************        
        loGrafico.DrawString("TOTAL: ", loFuenteTitulo, XBrushes.Black, loLinea, XStringFormats.TopLeft)

        lcTexto = lcMoneda & goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Net")), 2)
        loLineaDerecha.Width = loGrafico.MeasureString(lcTexto, loFuenteTitulo).Width
        loLineaDerecha.X = lnAnchoPagina - lnMargen - loLineaDerecha.Width
        loGrafico.DrawString(lcTexto, loFuenteTitulo, XBrushes.Black, loLineaDerecha, XStringFormats.TopLeft)

        loLinea.Offset(0, loLinea.Height)
        loLineaDerecha.Offset(0, loLineaDerecha.Height)

        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )
        loLinea.Offset(0, lnMargenC)
        loLineaDerecha.Offset(0, lnMargenC)
        loGrafico.DrawLine(loTrazo, lnMargen, loLinea.Y, lnAnchoPagina - lnMargen, loLinea.y )

        'Tamaño de la hoja
        loLinea.Offset(0, lnMargenV)

        Dim lnLongitud As Double = loLinea.Y

    '*****************************************************
    ' Cálculo del tamaño del papel
    '*****************************************************        
        Dim loAlturaInicial As Double = loPagina.Height.Point
        loPagina.Height = XUnit.FromPoint(lnLongitud)
        loPagina.MediaBox = New PdfRectangle(New XRect(0, loAlturaInicial - lnLongitud, loPagina.Width.Point, lnLongitud))

        Dim loSalida As New System.IO.MemoryStream()
        loDocumento.Save(loSalida)
        

    '*****************************************************
    ' Saca el PDF en pantalla
    '*****************************************************        
        Me.Response.Clear()
        Me.Response.Buffer = True

        Me.Response.ContentType = "application/pdf"
        'Me.Response.AddHeader("Content-Disposition", "attachment; filename=""Factura_" & lcFactura & ".pdf""")

        Me.Response.BinaryWrite(loSalida.GetBuffer())
        Me.Response.Flush()
        Me.Response.End()


    End Sub

    Private Sub mEscribirTexto(lcTexto As String)

        Me.Response.Clear()
        Me.Response.Buffer = True

        'Me.Response.ContentType = "application/pdf"
        Me.Response.ContentType = "text/plain"
        'Me.Response.AddHeader("Content-Disposition", "attachment; filename=""fFacturas_Ventas_TicketIpos.txt""")

        Me.Response.Write(lcTexto)
        Me.Response.Flush()
        Me.Response.End()
    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 21/02/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
