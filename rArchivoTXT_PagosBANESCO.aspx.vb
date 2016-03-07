'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_PagosBANESCO"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_PagosBANESCO
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
        
        Dim lcNombreEmpresa as String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcNombre)
        Dim lcRifEmpresa as String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)

        Dim lcOrden As String = goReportes.pcOrden
        
        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Proveedores.Cod_Pro                                             AS Cod_Pro,")
            loConsulta.AppendLine("        Proveedores.Nom_Pro                                             AS Nom_Pro,")
            loConsulta.AppendLine("        Proveedores.Rif                                                 AS Rif,")
            loConsulta.AppendLine("        COALESCE(   (")
            loConsulta.AppendLine("           SELECT  TOP 1 ")
            loConsulta.AppendLine("                    Cuentas_Clientes.Num_Cue ")
            loConsulta.AppendLine("            FROM    Cuentas_Clientes ")
            loConsulta.AppendLine("            WHERE   Cuentas_Clientes.Cod_Reg = Proveedores.Cod_Pro")
            loConsulta.AppendLine("                AND Cuentas_Clientes.Status = 'A'), '')                 AS Num_Cue,")
            loConsulta.AppendLine("        Montos.Num_Doc                                                  AS Referencia,")
            loConsulta.AppendLine("        Montos.Mon_Net                                                  AS Mon_Net,")
            loConsulta.AppendLine("        " & lcNombreEmpresa & "                                         AS Nombre_Empresa,")
            loConsulta.AppendLine("        " & lcRifEmpresa & "                                            AS Rif_Empresa,")
            loConsulta.AppendLine("        Cuentas_Bancarias.Num_Cue                                       AS Cuenta_Empresa,")
            loConsulta.AppendLine("        CAST( " & lcParametro5Desde & " AS DATE)                        AS Emision,")
            loConsulta.AppendLine("        CAST( " & lcParametro6Desde & " AS VARCHAR(8))                  AS Codigo_Descripcion_Pago,")
            loConsulta.AppendLine("        CAST( " & lcParametro7Desde & " AS VARCHAR(8))                  AS Referencia_Debito,")
            loConsulta.AppendLine("        CAST( " & lcParametro8Desde & " AS VARCHAR(12))                 AS Numero_Orden_Pago,")
            loConsulta.AppendLine("        CAST( " & lcParametro9Desde & " AS VARCHAR(2))                  AS Debitos_Individuales")
            loConsulta.AppendLine("FROM    Proveedores")
            loConsulta.AppendLine("    JOIN    Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = " & lcParametro4Desde)
            loConsulta.AppendLine("    JOIN (  SELECT      Pagos.Cod_Pro, ")
            loConsulta.AppendLine("                        Detalles_Pagos.Num_Doc,")
            loConsulta.AppendLine("                        Detalles_Pagos.Mon_Net ")
            loConsulta.AppendLine("            FROM        Pagos")
            loConsulta.AppendLine("                JOIN    Detalles_Pagos ")
            loConsulta.AppendLine("                    ON  Detalles_Pagos.Documento = Pagos.Documento")
            loConsulta.AppendLine("                    AND Tip_Ope = 'Transferencia'")
            loConsulta.AppendLine("            WHERE   Pagos.Status = 'Confirmado'")
            loConsulta.AppendLine("                AND Detalles_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                AND Pagos.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                AND " & lcParametro1Hasta)
            loConsulta.AppendLine("                AND Detalles_Pagos.Cod_Cue = " & lcParametro4Desde)
            loConsulta.AppendLine("            UNION ALL")
            loConsulta.AppendLine("            SELECT      Ordenes_Pagos.Cod_Pro, ")
            loConsulta.AppendLine("                        Detalles_oPagos.Num_Doc,")
            loConsulta.AppendLine("                        Detalles_oPagos.Mon_Net ")
            loConsulta.AppendLine("            FROM        Ordenes_Pagos")
            loConsulta.AppendLine("                JOIN    Detalles_oPagos ")
            loConsulta.AppendLine("                    ON  Detalles_oPagos.Documento = Ordenes_Pagos.Documento")
            loConsulta.AppendLine("                    AND Tip_Ope = 'Transferencia'")
            loConsulta.AppendLine("            WHERE       Ordenes_Pagos.Status = 'Confirmado'")
            loConsulta.AppendLine("                AND Detalles_oPagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                AND Ordenes_Pagos.Documento BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("                AND " & lcParametro2Hasta)
            loConsulta.AppendLine("                AND Detalles_oPagos.Cod_Cue = " & lcParametro4Desde)
            loConsulta.AppendLine("            ) Montos")
            loConsulta.AppendLine("        ON Montos.Cod_Pro = Proveedores.Cod_Pro")
            loConsulta.AppendLine("WHERE   Montos.Mon_Net > 0")
            loConsulta.AppendLine("    AND Proveedores.Cod_Pro BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("    AND " & lcParametro3Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrden)
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_PagosBANESCO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_PagosBANESCO.ReportSource = loObjetoReporte

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
    
    Private Sub mGenerarArchivoTxt(laDatosReporte As DataSet)
        Dim loTabla As DataTable = laDatosReporte.Tables(0)
        Dim loLimpiaCedula As New Regex("[^VEP0-9]")
        Dim loLimpiaRif As New Regex("[^VEJG0-9]")
        Dim loLimpiaNumero As New Regex("[^0-9]")
        Dim loLimpiaAlfaNum As New Regex("[^0-9a-zA-Z]")

        If (loTabla.Rows.Count = 0 ) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.rows(0)
        Dim ldFechaEmision As Date = CDate(loRenglon("Emision"))
        Dim lcNombreArchivo As String = "PAGOS_" & ldFechaEmision.ToString("ddMMyy")
        
        Dim llDebitosIndividuales As Boolean = (CStr(loRenglon("Debitos_Individuales")).Trim().ToUpper() = "SI")

        Dim loContenido As New StringBuilder()
        

        '**************************************************
        ' Encabezado fijo: identificación del tipo de TXT
        '**************************************************
        loContenido.AppendLine("HDRBANESCO        ED  95BPAYMULP")

        '**************************************************
        ' Encabezado variable 01: datos generales del TXT
        '**************************************************
        loContenido.Append("01")        'Tipo de renglon
        loContenido.Append("SCV")       'FIJO

        Dim lcCodigoDescripcionPago As String
        lcCodigoDescripcionPago =  CStr(loRenglon("Codigo_Descripcion_Pago"))
        lcCodigoDescripcionPago = loLimpiaNumero.Replace(lcCodigoDescripcionPago, "")
        lcCodigoDescripcionPago =  Strings.Left(lcCodigoDescripcionPago & Strings.Space(2), 2)
        loContenido.Append(lcCodigoDescripcionPago) 'Codigo de Descripcion del Pago

        loContenido.Append("                              9  ")       'FIJO

        Dim lcNumeroOrdenPago As String
        lcNumeroOrdenPago =  CStr(loRenglon("Numero_Orden_Pago")).Trim().ToUpper()
        lcNumeroOrdenPago = loLimpiaAlfaNum.Replace(lcNumeroOrdenPago, "")
        lcNumeroOrdenPago =  Strings.Left(lcNumeroOrdenPago & Strings.Space(12), 12)
        loContenido.Append(lcNumeroOrdenPago)           'Numero de la Orden de Pago (manual/opcional)

        loContenido.Append("                       ")   'FIJO

        'Encabezado: Fecha+hora actual (automático)
        Dim lcFechaHora As String
        Dim ldFechaHora As Date = Date.Now()
        lcFechaHora = ldFechaHora.ToString("yyyyMMddHHmmss")
        loContenido.Append(lcFechaHora)

        'Encabezado: Fin encabezado 01
        loContenido.AppendLine()

        'Monto Total (se obtiene de los renglones).
        Dim lnMontoTotal As Long
        For Each loItem As DataRow In loTabla.Rows
            lnMontoTotal += CLng(CDec(loItem("Mon_Net"))*100)
        Next
        Dim lcMontoTotal As String = lnMontoTotal.ToString("000000000000000")

        '**************************************************
        ' Encabezado variable 02: datos generales del TXT
        ' NOTA: Si los debitos son individuales se crea un 
        ' "Encabezado 02" para cada crédito, si no se crea
        ' un solo "Encabezado 02" global
        '**************************************************
        If Not llDebitosIndividuales Then

            loContenido.Append("02")        'Tipo de renglon
            
            'Encabezado: Referencia del débito (manual/opcional)
            Dim lcReferenciaDelDebito As String

            lcReferenciaDelDebito =  CStr(loRenglon("Referencia_Debito"))
            lcReferenciaDelDebito = loLimpiaNumero.Replace(lcReferenciaDelDebito, "")
            If (lcReferenciaDelDebito = "") Then lcReferenciaDelDebito = "5"
            lcReferenciaDelDebito =  Strings.Right(Strings.StrDup(8, "0") & lcReferenciaDelDebito, 8)
            loContenido.Append(lcReferenciaDelDebito)

            'Encabezado: FIJO.
            loContenido.Append("                      ")

            'Encabezado: RIF del Ordenante (se toma de la empresa actual).
            Dim lcRifOrdenante As String
            lcRifOrdenante =  CStr(loRenglon("Rif_Empresa")).Trim()
            lcRifOrdenante = loLimpiaRif.Replace(lcRifOrdenante, "")
            lcRifOrdenante =  Strings.Left(lcRifOrdenante & Strings.Space(17), 17)
            loContenido.Append(lcRifOrdenante)

            'Encabezado: Nombre del Ordenante (se toma de la empresa actual). 
            Dim lcNombreOrdenante As String
            lcNombreOrdenante = CStr(loRenglon("Nombre_Empresa")).Trim()
            lcNombreOrdenante = Me.mEliminarAcentos(lcNombreOrdenante)
            lcNombreOrdenante = Strings.Left(lcNombreOrdenante & Strings.Space(35), 35)
            loContenido.Append(lcNombreOrdenante)

            'Encabezado: Monto Total (se obtiene de los renglones).
            loContenido.Append(lcMontoTotal)

            'Encabezado: Moneda (fijo, 4 caracteres)
            loContenido.Append("VEF ")

            'Encabezado: Numero de cuenta de la empresa
            Dim lcNumeroCuentaDebitar As String = CStr(loRenglon("Cuenta_Empresa")).Trim()
            lcNumeroCuentaDebitar = loLimpiaNumero.Replace(lcNumeroCuentaDebitar, "")
            lcNumeroCuentaDebitar = Strings.Left(lcNumeroCuentaDebitar & Strings.Space(20), 20)
            loContenido.Append(lcNumeroCuentaDebitar)

            'Encabezado: FIJO
            loContenido.Append("              BANESCO    ")    

            'Encabezado: Fecha de ejecución del pago
            Dim lcFechaPago As String = ldFechaEmision.ToString("yyyyMMdd")
            loContenido.Append(lcFechaPago)
            loContenido.AppendLine()

        End If

        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        Dim lnCantidad As Integer = loTabla.Rows.Count
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            Dim lnMonto As Long 
            Dim lcMonto As String 


            If llDebitosIndividuales Then

                    '**************************************************
                    ' Encabezado variable 02: datos generales del TXT
                    ' (Para registrar Debitos Individuales)
                    '**************************************************
                    loContenido.Append("02")        'Tipo de renglon
                    
                    'Encabezado: Referencia del débito (manual/opcional)
                    Dim lcReferenciaDelDebito As String

                    lcReferenciaDelDebito =  (n+1).ToString("00000000")
                    loContenido.Append(lcReferenciaDelDebito)

                    'Encabezado: FIJO.
                    loContenido.Append("                      ")

                    'Encabezado: RIF del Ordenante (se toma de la empresa actual).
                    Dim lcRifOrdenante As String
                    lcRifOrdenante =  CStr(loRenglon("Rif_Empresa")).Trim()
                    lcRifOrdenante = loLimpiaRif.Replace(lcRifOrdenante, "")
                    lcRifOrdenante =  Strings.Left(lcRifOrdenante & Strings.Space(17), 17)
                    loContenido.Append(lcRifOrdenante)

                    'Encabezado: Nombre del Ordenante (se toma de la empresa actual). 
                    Dim lcNombreOrdenante As String
                    lcNombreOrdenante = CStr(loRenglon("Nombre_Empresa")).Trim()
                    lcNombreOrdenante = Me.mEliminarAcentos(lcNombreOrdenante)
                    lcNombreOrdenante = Strings.Left(lcNombreOrdenante & Strings.Space(35), 35)
                    loContenido.Append(lcNombreOrdenante)

                    'Encabezado: Monto Individual (15 caracteres, los dos últimos son decimales, rellenar con "0" a la izq.) 
                    lnMonto = CLng(CDec(loRenglon("Mon_Net"))*100)
                    lcMonto = lnMonto.ToString("000000000000000")
                    loContenido.Append(lcMonto)

                    'Encabezado: Moneda (fijo, 4 caracteres)
                    loContenido.Append("VEF ")

                    'Encabezado: Numero de cuenta de la empresa
                    Dim lcNumeroCuentaDebitar As String = CStr(loRenglon("Cuenta_Empresa")).Trim()
                    lcNumeroCuentaDebitar = loLimpiaNumero.Replace(lcNumeroCuentaDebitar, "")
                    lcNumeroCuentaDebitar = Strings.Left(lcNumeroCuentaDebitar & Strings.Space(20), 20)
                    loContenido.Append(lcNumeroCuentaDebitar)

                    'Encabezado: FIJO
                    loContenido.Append("              BANESCO    ")    

                    'Encabezado: Fecha de ejecución del pago
                    Dim lcFechaPago As String = ldFechaEmision.ToString("yyyyMMdd")
                    loContenido.Append(lcFechaPago)
                    loContenido.AppendLine()

            End If

            'Datos: Tipo de renglon (FIJO)
            loContenido.Append("03")

            'Datos: Referencia del crédito (30 caracteres, comienza formateado con 8 ceros, rellenar con espacios)
            Dim lcReferenciaCredito As String = (n+1).ToString("00000000")
            lcReferenciaCredito = Strings.Left(lcReferenciaCredito & Strings.Space(30), 30)
            loContenido.Append(lcReferenciaCredito)

            'Datos: Monto Individual (15 caracteres, los dos últimos son decimales, rellenar con "0" a la izq.)
            lnMonto = CLng(CDec(loRenglon("Mon_Net"))*100)
            lcMonto = lnMonto.ToString("000000000000000")
            loContenido.Append(lcMonto)

            'Datos: Moneda (fijo, 3 caracteres))
            loContenido.Append("VEF")

            'Datos: Cuenta (20 caracteres, rellenar con X en caso de error)
            Dim lcCuenta As String = CStr(loRenglon("Num_Cue")).Trim()
            lcCuenta = Strings.Left(lcCuenta & "XXXXXXXXXXXXXXXXXXXX", 20)
            loContenido.Append(lcCuenta)

            'Datos: FIJO (10 espacios)
            loContenido.Append("          ")

            'Datos: Código SUDEBAN del banco/Cuenta (primeros 4 caracteres de la cuenta, rellenar con X en caso de error)
            lcCuenta = Strings.Left(lcCuenta & "XXXX", 4)
            loContenido.Append(lcCuenta)

            'Datos: FIJO (10 espacios)
            loContenido.Append("          ")

            'Datos: Cédula (17 caracteres, rellenar con espacios a la derecha)
            Dim lcRif As String = CStr(loRenglon("Rif")).ToUpper()
            lcRif = loLimpiaRif.Replace(lcRif, "")
            If (lcRif = loLimpiaNumero.Replace(lcRif, "") )
                lcRif = "X" & lcRif
            End If
            lcRif = Strings.Left(lcRif & Strings.Space(17), 17)
            loContenido.Append(lcRif)

            'Datos: Nombre del trabajador (35 caracteres, rellenar con espacios)
            Dim lcNombre As String = CStr(loRenglon("Nom_Pro")).ToUpper()
            lcNombre = Me.mEliminarAcentos(lcNombre)
            lcNombre = Strings.Left(lcNombre & Strings.Space(35), 35)
            loContenido.Append(lcNombre)

            'Datos: FIJO (236 espacios)
            loContenido.Append(Strings.Space(236))

            'Datos: FIJO 
            loContenido.Append("42")

            'Datos: Cuenta Banesco (" ") u Otra ("5")
            If (lcCuenta = "0134") Then
                loContenido.Append(" ")
            Else
                loContenido.Append("5")
            End If
            
            'Fin de línea de detalle
            loContenido.Append(vbNewLine)
        
        Next n
        
        'Pie: Tipo de renglon
        loContenido.Append("06")

        'Pie: Cantidad de Débitos (15 dígitos, relleno con ceros a la izquierda)
        If llDebitosIndividuales Then 
            loContenido.Append(loTabla.Rows.Count.ToString("000000000000000"))
        Else
            loContenido.Append(1.ToString("000000000000000"))
        End If

        'Pie: Cantidad de Créditos (15 dígitos, relleno con ceros a la izquierda)
        loContenido.Append(loTabla.Rows.Count.ToString("000000000000000"))

        'Pie: Monto Total (15 dígitos, relleno con ceros a la izquierda)
        loContenido.Append(lcMontoTotal)

        'Pie: Fin de línea/archivo
        loContenido.Append(vbNewLine)



        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcNombreArchivo & ".txt")
        Me.Response.ContentType = "text/plain"
        Me.Response.Write(loContenido.ToString())
        Me.Response.End()

    End Sub

    Private Function mEliminarAcentos(lcTexto As String ) As String 

        lcTexto = Regex.Replace(lcTexto, "[áàâä]", "a")
        lcTexto = Regex.Replace(lcTexto, "[éèêë]", "e")
        lcTexto = Regex.Replace(lcTexto, "[íìîï]", "i")
        lcTexto = Regex.Replace(lcTexto, "[óòôö]", "o")
        lcTexto = Regex.Replace(lcTexto, "[úùûü]", "u")
        lcTexto = Regex.Replace(lcTexto, "ñ", "n")
        lcTexto = Regex.Replace(lcTexto, "ç", "c")

        lcTexto = Regex.Replace(lcTexto, "[ÁÀÂÄ]", "A")
        lcTexto = Regex.Replace(lcTexto, "[ÉÈÊË]", "E")
        lcTexto = Regex.Replace(lcTexto, "[ÍÌÎÏ]", "I")
        lcTexto = Regex.Replace(lcTexto, "[ÓÒÔÖ]", "O")
        lcTexto = Regex.Replace(lcTexto, "[ÚÙÛÜ]", "U")
        lcTexto = Regex.Replace(lcTexto, "Ñ", "N")
        lcTexto = Regex.Replace(lcTexto, "Ç", "C")
        
        lcTexto = regex.Replace(lcTexto, "[^0-9a-zA-Z ]", "")

        Return lcTexto
    End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 24/04/15: Código Inicial.
'-------------------------------------------------------------------------------------------'
