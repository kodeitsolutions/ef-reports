'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_KartuvalIpos"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_KartuvalIpos
    Inherits vis2formularios.frmReporte

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        loConsulta.AppendLine("")
        loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
        loConsulta.AppendLine("DECLARE @lcFormaPago VARCHAR(MAX);")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("SET @lcDocumento = (")
        loConsulta.AppendLine("    SELECT TOP 1 Documento")
        loConsulta.AppendLine("    FROM Facturas ")
        loConsulta.AppendLine("    WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal & ");")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("-- **************************************************************************")
        loConsulta.AppendLine("--  Busca las formas de pago del cobro de la factura")
        loConsulta.AppendLine("--  y las concatena en una variable.")
        loConsulta.AppendLine("-- **************************************************************************")
        loConsulta.AppendLine("SELECT      @lcFormaPago = COALESCE(@lcFormaPago + ',',  '') + ")
        loConsulta.AppendLine("           RTRIM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta'")
        loConsulta.AppendLine("                THEN 'Tarjeta' + (CASE Tarjetas.Tip_Tar ")
        loConsulta.AppendLine("                                    WHEN 'C' THEN '/Crédito' ")
        loConsulta.AppendLine("                                    WHEN 'D' THEN '/Débito' ")
        loConsulta.AppendLine("                                    ELSE '/X' END )")
        loConsulta.AppendLine("                ELSE Detalles_Cobros.Tip_Ope")
        loConsulta.AppendLine("            END)")
        loConsulta.AppendLine("FROM        Detalles_Cobros")
        loConsulta.AppendLine("    JOIN    Renglones_Cobros ")
        loConsulta.AppendLine("        ON  Renglones_Cobros.Documento = Detalles_Cobros.Documento")
        loConsulta.AppendLine("        AND Renglones_Cobros.Cod_Tip = 'FACT'")
        loConsulta.AppendLine("        AND Renglones_Cobros.Doc_Ori = @lcDocumento")
        loConsulta.AppendLine("    LEFT JOIN Tarjetas")
        loConsulta.AppendLine("        ON  Tarjetas.Cod_Tar = Detalles_Cobros.Cod_Tar; ")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("")
        loConsulta.AppendLine("-- **************************************************************************")
        loConsulta.AppendLine("--  SELECT General: recupera todos lso datos de la factura y sus renglones")
        loConsulta.AppendLine("-- **************************************************************************")
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
        loConsulta.AppendLine("            Facturas.Mon_Exe                            AS Mon_Exe, ")
        loConsulta.AppendLine("            Facturas.Mon_Bru - Facturas.Mon_Exe         AS Mon_Bas, ")
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
        loConsulta.AppendLine("            Renglones_Facturas.Mon_Imp1                 AS Impuesto,")
        loConsulta.AppendLine("            @lcFormaPago                                AS Formas_de_Pago")
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
        Dim lcFormaDePago As String = CStr(loFactura("Formas_de_Pago")).Trim()
        Dim lcComentario As String = CStr(loFactura("Comentario")).Trim()

        Dim lcCliente_Rif As String = CStr(loFactura("Rif")).Trim()
        Dim lcCliente_Nit As String = CStr(loFactura("Nit")).Trim()
        Dim lcCliente_Codigo As String = CStr(loFactura("Cod_Cli")).Trim()
        Dim lcCliente_Nombre As String = CStr(loFactura("Nom_Cli")).Trim()
        Dim lcCliente_Direccion As String = CStr(loFactura("Dir_Fis")).Trim()
        Dim lcCliente_Telefono As String = CStr(loFactura("Telefonos")).Trim()

    '*****************************************************
    ' Definición del PDF (documento y valores)
    '*****************************************************
        Dim loDocumento As New PdfDocument()
        Dim loPagina As PdfPage = loDocumento.AddPage()
        Dim loGrafico As XGraphics = XGraphics.FromPdfPage(loPagina)

        'Define el ancho y alto de la página (aproximadamente 1/1 carta: 8,5in x 5,9in)
        loPagina.Width = XUnit.FromInch(8.5)
        loPagina.Height = XUnit.FromInch(5.9)

        Dim lnAnchoPagina As Double = loPagina.Width.Point
        Dim lnInternineado As Double = XUnit.FromMillimeter(2)

        Dim lnMargenSup As Double = XUnit.FromMillimeter(27)
        Dim lnMargenH As Double = XUnit.FromMillimeter(10)
        Dim lnMargenC As Double = XUnit.FromMillimeter(1)

        Dim loFuenteTitulo As XFont = new XFont("Arial", 6, XFontStyle.Bold)
        Dim loFuente As XFont = new XFont("Arial", 8, XFontStyle.Regular)
        Dim loFuenteN As XFont = new XFont("Arial", 8, XFontStyle.Bold)
        
        Dim loTrazo As New XPen(XColor.FromArgb(0), 0.5) 'Para dibujar línea horizontal de separación

    '*****************************************************
    ' Encabezado
    '*****************************************************
        Dim loPosicion As XPoint
        loPosicion.X = lnMargenH
        loPosicion.Y = lnMargenSup
        
        
        Dim laDireccion() As String  = goEmpresa.pcDireccionEmpresa.Split(New String(){vbNewLine, vbLf, vbCr},  StringSplitOptions.RemoveEmptyEntries)
        For Each lcLinea As String In laDireccion
            loGrafico.DrawString(lcLinea, loFuenteTitulo, XBrushes.Black, loPosicion)
            loPosicion.Y += loFuenteTitulo.Height
        Next lcLinea

        loGrafico.DrawString("Telf: " & goEmpresa.pcTelefonoEmpresa & " E-mail: " & goEmpresa.pcCorreo, _
                             loFuenteTitulo, XBrushes.Black, loPosicion)

        loPosicion.X = XUnit.FromMillimeter(150).Point
        'loPosicion.Y += loFuenteTitulo.Height
        loGrafico.DrawString("SUNACOP Nro: 384112", loFuenteN, XBrushes.Black, loPosicion)

        loPosicion.Y += loFuenteN.Height
        loGrafico.DrawString("FACTURA Nro: " & lcFactura, loFuenteN, XBrushes.Black, loPosicion)

        loPosicion.X = XUnit.FromMillimeter(130).Point
        loGrafico.DrawString("CONTADO", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.Y += loFuenteN.Height


    '*****************************************************
    ' Datos del cliente
    '*****************************************************
        loPosicion.X = lnMargenH
        
        loGrafico.DrawString("Código: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += XUnit.FromMillimeter(18).Point
        loGrafico.DrawString(lcCliente_Codigo, loFuente, XBrushes.Black, loPosicion)
        
        loPosicion.X += XUnit.FromMillimeter(20).Point
        loGrafico.DrawString("R.I.F.: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += loGrafico.MeasureString("R.I.F._", loFuenteN).Width
        loGrafico.DrawString(lcCliente_Rif, loFuente, XBrushes.Black, loPosicion)
        
        loPosicion.X += XUnit.FromMillimeter(30).Point
        loGrafico.DrawString("N.I.T.: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += loGrafico.MeasureString("N.I.T._", loFuenteN).Width
        loGrafico.DrawString(lcCliente_Nit, loFuente, XBrushes.Black, loPosicion)
        
        loPosicion.X = lnMargenH
        loPosicion.Y += loFuenteN.Height
        loGrafico.DrawString("Cliente: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += XUnit.FromMillimeter(18).Point
        loGrafico.DrawString(lcCliente_Nombre, loFuente, XBrushes.Black, loPosicion)

        loPosicion.X = lnMargenH
        loPosicion.Y += loFuenteN.Height
        loGrafico.DrawString("Direccción: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += XUnit.FromMillimeter(18).Point

        'La dirección se divide hasta en 2 líneas
        Dim lcDireccion1 As String = lcCliente_Direccion 
        Dim lcDireccion2 As String = ""
        Dim lnAnchoDireccion As Double = XUnit.FromMillimeter(150).Point - loPosicion.X
        While(loGrafico.MeasureString(lcDireccion1, loFuente).Width > lnAnchoDireccion)
            lcDireccion2 = lcDireccion1.Substring(lcDireccion1.Length-1, 1) & lcDireccion2
            lcDireccion1 = lcDireccion1.Substring(0, lcDireccion1.Length-1)
        End While
        While(loGrafico.MeasureString(lcDireccion2, loFuente).Width > lnAnchoDireccion)
            lcDireccion2 = lcDireccion2.Substring(0, lcDireccion2.Length-1)
        End While
        
        loGrafico.DrawString(lcDireccion1, loFuente, XBrushes.Black, loPosicion)
        loPosicion.Y += loFuente.Height
        loGrafico.DrawString(lcDireccion2, loFuente, XBrushes.Black, loPosicion)
        
        loPosicion.X = lnMargenH
        loPosicion.Y += loFuenteN.Height
        loGrafico.DrawString("Telf(s): ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += XUnit.FromMillimeter(18).Point
        loGrafico.DrawString(lcCliente_Telefono, loFuente, XBrushes.Black, loPosicion)
        
        loPosicion.X = XUnit.FromMillimeter(150).Point
        loGrafico.DrawString("Valencia, " & Date.Now().ToString("dd/MM/yyyy"), loFuenteN, XBrushes.Black, loPosicion)
        
    '*****************************************************
    ' Encabezado de Detalle
    '*****************************************************
        loPosicion.X = lnMargenH
        loPosicion.Y += loFuenteN.Height
        Dim lnAlturaCaja As Double = lnMargenC*2 + loFuenteN.Height
        Dim lnBarra_1 As Double = lnMargenH
        Dim lnBarra_2 As Double = lnBarra_1 + XUnit.FromMillimeter(14).Point
        Dim lnBarra_3 As Double = lnBarra_2 + XUnit.FromMillimeter(100).Point
        Dim lnBarra_4 As Double = lnBarra_3 + XUnit.FromMillimeter(20).Point
        Dim lnBarra_5 As Double = lnBarra_4 + XUnit.FromMillimeter(25).Point
        Dim lnBarra_6 As Double = lnBarra_5 + XUnit.FromMillimeter(25).Point
        loGrafico.DrawLine(loTrazo, lnBarra_1, loPosicion.Y, lnBarra_6, loPosicion.y )
        loGrafico.DrawLine(loTrazo, lnBarra_1, loPosicion.Y, lnBarra_1, loPosicion.Y+lnAlturaCaja )
        loGrafico.DrawLine(loTrazo, lnBarra_2, loPosicion.Y, lnBarra_2, loPosicion.Y+lnAlturaCaja )
        loGrafico.DrawLine(loTrazo, lnBarra_3, loPosicion.Y, lnBarra_3, loPosicion.Y+lnAlturaCaja )
        loGrafico.DrawLine(loTrazo, lnBarra_4, loPosicion.Y, lnBarra_4, loPosicion.Y+lnAlturaCaja )
        loGrafico.DrawLine(loTrazo, lnBarra_5, loPosicion.Y, lnBarra_5, loPosicion.Y+lnAlturaCaja )
        loGrafico.DrawLine(loTrazo, lnBarra_6, loPosicion.Y, lnBarra_6, loPosicion.Y+lnAlturaCaja )
        
        loPosicion.Y += lnMargenC

        loGrafico.DrawString("Código", loFuenteN, XBrushes.Black, New XRect(lnBarra_1, loPosicion.Y, lnBarra_2 - lnBarra_1, loPosicion.Y) , XStringFormats.TopCenter )
        loGrafico.DrawString("Descripción", loFuenteN, XBrushes.Black, New XRect(lnBarra_2, loPosicion.Y, lnBarra_3 - lnBarra_2, loPosicion.Y) , XStringFormats.TopCenter )
        loGrafico.DrawString("Cantidad", loFuenteN, XBrushes.Black, New XRect(lnBarra_3, loPosicion.Y, lnBarra_4 - lnBarra_3, loPosicion.Y) , XStringFormats.TopCenter )
        loGrafico.DrawString("Precio", loFuenteN, XBrushes.Black, New XRect(lnBarra_4, loPosicion.Y, lnBarra_5 - lnBarra_4, loPosicion.Y) , XStringFormats.TopCenter )
        loGrafico.DrawString("Total", loFuenteN, XBrushes.Black, New XRect(lnBarra_5, loPosicion.Y, lnBarra_6 - lnBarra_5, loPosicion.Y) , XStringFormats.TopCenter )

        loPosicion.Y += loFuenteN.Height

        loPosicion.X = lnBarra_1
        loPosicion.Y += lnMargenC
        loGrafico.DrawLine(loTrazo, lnBarra_1, loPosicion.Y, lnBarra_6, loPosicion.y )



    '*****************************************************
    ' Artículos
    '*****************************************************
        loPosicion.Y += loFuenteN.Height + lnMargenC

        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        For Each loRenglon As DataRow In loRenglones.Rows
            Dim lcCodigo As String = CStr(loRenglon("Cod_Art"))
            Dim lcNombre As String = CStr(loRenglon("Nom_Art"))
            Dim lcCantidad As String = goServicios.mObtenerFormatoCadena(IIf(lnDecimalesCantidad>0, CDec(loRenglon("Can_Art1")), CInt(Decimal.Floor(loRenglon("Can_Art1")))), lnDecimalesCantidad)
            Dim lcPrecio As String = goServicios.mObtenerFormatoCadena(CDec(loRenglon("Precio1")), lnDecimalesMonto)
            Dim lcTotal As String = goServicios.mObtenerFormatoCadena(CDec(loRenglon("Neto")), lnDecimalesMonto)
            
            loPosicion.X = lnBarra_1 + XUnit.FromMillimeter(3).Point
            loGrafico.DrawString(lcCodigo, loFuente, XBrushes.Black, loPosicion)

            loPosicion.X = lnBarra_2 + XUnit.FromMillimeter(3).Point
            loGrafico.DrawString(lcNombre, loFuente, XBrushes.Black, loPosicion)
            
            loPosicion.X = lnBarra_4 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcCantidad, loFuente).Width
            loGrafico.DrawString(lcCantidad, loFuente, XBrushes.Black, loPosicion)

            loPosicion.X = lnBarra_5 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcPrecio, loFuente).Width
            loGrafico.DrawString(lcPrecio, loFuente, XBrushes.Black, loPosicion)

            loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcTotal, loFuente).Width
            loGrafico.DrawString(lcTotal, loFuente, XBrushes.Black, loPosicion)

            loPosicion.Y += loFuente.Height

        Next loRenglon

        loPosicion.Y = lnMargenSup + XUnit.FromMillimeter(90).Point
        loPosicion.X = lnBarra_1
        loGrafico.DrawLine(loTrazo, lnBarra_1, loPosicion.Y, lnBarra_6, loPosicion.y )

        
    '*****************************************************
    ' Forma de pago y observaciones
    '*****************************************************

        Dim lnInicioTotales As Double = loPosicion.Y + loFuenteN.Height

        loPosicion.X = lnMargenH
        loPosicion.Y = lnInicioTotales

        loGrafico.DrawString("Forma de Pago: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += loGrafico.MeasureString("Observaciones:_", loFuenteN).Width
        loGrafico.DrawString(lcFormaDePago, loFuente, XBrushes.Black, loPosicion)

        loPosicion.X = lnMargenH
        loPosicion.Y += loFuenteN.Height
        
        loGrafico.DrawString("Observaciones: ", loFuenteN, XBrushes.Black, loPosicion)
        loPosicion.X += loGrafico.MeasureString("Observaciones:_", loFuenteN).Width

        'El comentario/observaciones se divide hasta en 3 líneas
        Dim lcComentario1 As String = lcComentario
        Dim lcComentario2 As String = ""
        Dim lcComentario3 As String = ""
        Dim lnAnchoComentario As Double = XUnit.FromMillimeter(140).Point - loPosicion.X

        While(loGrafico.MeasureString(lcComentario1, loFuente).Width > lnAnchoComentario)
            lcComentario2 = lcComentario1.Substring(lcComentario1.Length-1, 1) & lcComentario2
            lcComentario1 = lcComentario1.Substring(0, lcComentario1.Length-1)
        End While
        While(loGrafico.MeasureString(lcComentario2, loFuente).Width > lnAnchoComentario)
            lcComentario3 = lcComentario2.Substring(lcComentario2.Length-1, 1) & lcComentario3
            lcComentario2 = lcComentario2.Substring(0, lcComentario2.Length-1)
        End While
        While(loGrafico.MeasureString(lcComentario3, loFuente).Width > lnAnchoComentario)
            lcComentario3 = lcComentario3.Substring(0, lcComentario3.Length-1)
        End While
        
        loGrafico.DrawString(lcComentario1, loFuente, XBrushes.Black, loPosicion)
        loPosicion.Y += loFuente.Height
        loGrafico.DrawString(lcComentario2, loFuente, XBrushes.Black, loPosicion)
        loPosicion.Y += loFuente.Height
        loGrafico.DrawString(lcComentario3, loFuente, XBrushes.Black, loPosicion)

        loPosicion.Y += loFuenteN.Height
        
    '*****************************************************
    ' Sub total, Descuentos y Recargos
    '*****************************************************
        loPosicion.X = lnMargenH  + XUnit.FromMillimeter(135).Point
        loPosicion.Y = lnInicioTotales
        Dim lcMonto As String = ""

        lcMonto = goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Bru")), lnDecimalesMonto)
        loGrafico.DrawString("SubTotal: ", loFuenteN, XBrushes.Black, loPosicion)

        loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcMonto, loFuenteN).Width
        loGrafico.DrawString(lcMonto, loFuenteN, XBrushes.Black, loPosicion)


        loPosicion.X = lnMargenH  + XUnit.FromMillimeter(135).Point
        loPosicion.Y += loFuenteN.Height

        lcMonto = goServicios.mObtenerFormatoCadena(-CDec(loFactura("Mon_Des")), lnDecimalesMonto)
        loGrafico.DrawString("Descuento: ", loFuente, XBrushes.Black, loPosicion)

        loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcMonto, loFuente).Width
        loGrafico.DrawString(lcMonto, loFuente, XBrushes.Black, loPosicion)

        
        loPosicion.X = lnMargenH  + XUnit.FromMillimeter(135).Point
        loPosicion.Y += loFuenteN.Height

        lcMonto = goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Bas")), lnDecimalesMonto)
        loGrafico.DrawString("Base Imponible: ", loFuente, XBrushes.Black, loPosicion)

        loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcMonto, loFuente).Width
        loGrafico.DrawString(lcMonto, loFuente, XBrushes.Black, loPosicion)


        loPosicion.X = lnMargenH  + XUnit.FromMillimeter(135).Point
        loPosicion.Y += loFuenteN.Height

        lcMonto = goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Imp1")), lnDecimalesMonto)
        loGrafico.DrawString("I.V.A: ", loFuenteN, XBrushes.Black, loPosicion)

        loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcMonto, loFuenteN).Width
        loGrafico.DrawString(lcMonto, loFuenteN, XBrushes.Black, loPosicion)
        
        loPosicion.X = lnMargenH + XUnit.FromMillimeter(135).Point - XUnit.FromMillimeter(2).Point
        loPosicion.Y += lnMargenC
        loGrafico.DrawLine(loTrazo, loPosicion.X, loPosicion.Y, lnBarra_6, loPosicion.y )
        loPosicion.Y += lnMargenC

        loPosicion.X = lnMargenH  + XUnit.FromMillimeter(135).Point
        loPosicion.Y += loFuenteN.Height

        lcMonto = goServicios.mObtenerFormatoCadena(CDec(loFactura("Mon_Net")), lnDecimalesMonto)
        loGrafico.DrawString("Total Facturado: ", loFuenteN, XBrushes.Black, loPosicion)

        loPosicion.X = lnBarra_6 - XUnit.FromMillimeter(6).Point - loGrafico.MeasureString(lcMonto, loFuenteN).Width
        loGrafico.DrawString(lcMonto, loFuenteN, XBrushes.Black, loPosicion)





    '*****************************************************
    '*****************************************************
    '*****************************************************
    '*****************************************************


    '*****************************************************
    ' "Compila" el PDF (en memoria)
    '*****************************************************        
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
' RJG: 24/02/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
