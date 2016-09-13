'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_Nomina_100BANCO"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_Nomina_100BANCO
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))'Recibo
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))'Trabajador
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))'Contrato del trabajador
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) ' Descripción
        
        Dim lcCodigoEmpresa as String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)
        Dim lcNombreEmpresa as String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcNombre)
        Dim lcRifEmpresa as String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)

        Dim lcOrden As String = goReportes.pcOrden
        
        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                            AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                            AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula                             AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Num_Cue                            AS Num_Cue,")
            loConsulta.AppendLine("            ROUND(Pagos.Mon_Net, 2)                         AS Mon_Net,")
            loConsulta.AppendLine("            Datos_Codigo_Empresa.Val_Car                    AS Codigo_Empresa,")
            loConsulta.AppendLine("            " & lcNombreEmpresa & "                         AS Nombre_Empresa,")
            loConsulta.AppendLine("            " & lcRifEmpresa & "                            AS Rif_Empresa,")
            'loConsulta.AppendLine("            Cuentas_Bancarias.Num_Cue                       AS Cuenta_Debitar,")
            loConsulta.AppendLine("            COALESCE(CB_Empresa.Num_Cue, '')                 AS Cuenta_Empresa,")
            loConsulta.AppendLine("            CAST( " & lcParametro4Desde & " AS VARCHAR(MAX)) AS Descripcion")
            loConsulta.AppendLine("FROM        Trabajadores")
            loConsulta.AppendLine("    JOIN Campos_Propiedades Datos_Codigo_Empresa")
            loConsulta.AppendLine("         ON  Datos_Codigo_Empresa.Cod_Pro =  'CODEMP100%' ")
            loConsulta.AppendLine("         AND Datos_Codigo_Empresa.Cod_Reg =  " & lcCodigoEmpresa)
            loConsulta.AppendLine("         AND Datos_Codigo_Empresa.Cod_Reg =  " & lcCodigoEmpresa)
            loConsulta.AppendLine("         AND Datos_Codigo_Empresa.Origen = 'empresas' ")
            loConsulta.AppendLine("    JOIN Campos_Propiedades Datos_Cuenta_Empresa")
            loConsulta.AppendLine("         ON  Datos_Cuenta_Empresa.Cod_Pro =  'CODCUE100%' ")
            loConsulta.AppendLine("         AND Datos_Cuenta_Empresa.Cod_Reg =  " & lcCodigoEmpresa)
            loConsulta.AppendLine("         AND Datos_Cuenta_Empresa.Cod_Reg =  " & lcCodigoEmpresa)
            loConsulta.AppendLine("         AND Datos_Cuenta_Empresa.Origen = 'empresas' ")
            loConsulta.AppendLine("    JOIN Cuentas_Bancarias CB_Empresa ")
            loConsulta.AppendLine("         ON  CB_Empresa.Cod_Cue = Datos_Cuenta_Empresa.Val_Car")
            loConsulta.AppendLine("         AND CB_Empresa.Cod_Cue = Trabajadores.Cod_Cue ")
            loConsulta.AppendLine("    JOIN  ( SELECT  SUM(Recibos.Mon_Net) AS Mon_Net,")
            loConsulta.AppendLine("                    Recibos.Cod_Tra")
            loConsulta.AppendLine("            FROM    Recibos")
            loConsulta.AppendLine("            WHERE   Recibos.Cod_Con NOT IN  ('92','93','94','95')")
            loConsulta.AppendLine("                AND Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("                AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                AND Recibos.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                AND " & lcParametro1Hasta)
            loConsulta.AppendLine("            GROUP BY Recibos.Cod_Tra")
            loConsulta.AppendLine("            ) AS Pagos")
            loConsulta.AppendLine("        ON  Pagos.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("WHERE   Pagos.Mon_Net > 0")
            loConsulta.AppendLine("    AND Trabajadores.Tip_Pag = 'Transferencia'")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Tra BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Trabajadores.Cod_Con BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("    AND " & lcParametro3Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrden)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString() )
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_Nomina_100BANCO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_Nomina_100BANCO.ReportSource = loObjetoReporte

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
        Dim ldFechaGeneracion AS Date = Date.Now()
        Dim lcNombreArchivo As String = "NOMINA_" & ldFechaGeneracion.ToString("ddMMyy")

        Dim loContenido As New StringBuilder()
        
        
        '**************************************************
        ' Encabezado fijo: Concecutivo
        '**************************************************
        loContenido.Append("000000")
        '**************************************************
        ' Encabezado fijo: fecha + hora creacion
        '**************************************************
        loContenido.Append(ldFechaGeneracion.ToString("yyyyMMddhhmmss"))
        '**************************************************
        ' Encabezado fijo: fecha + hora efectiva
        '**************************************************
        loContenido.Append(ldFechaGeneracion.ToString("yyyyMMddhhmmss"))
        '**************************************************
        ' Encabezado fijo: fecha + hora aplicacion
        '**************************************************
        loContenido.Append(ldFechaGeneracion.ToString("yyyyMMddhhmmss"))

        '**************************************************
        ' Encabezado: código de la empresa
        '**************************************************
        loContenido.Append(Strings.Right("000000" & CStr(loRenglon("Codigo_Empresa")).Trim(), 6))

        '**************************************************
        ' Encabezado fijo: código de servicio
        '**************************************************
        loContenido.Append("000078")

        '**************************************************
        ' Encabezado: Tipo de cuenta + Numero(débito)
        '**************************************************
        loContenido.Append("CC ")
        loContenido.Append(Strings.Right("0000000000000000000000" & CStr(loRenglon("Cuenta_Empresa")).Trim(), 22))

        '**************************************************
        ' Encabezado fijo: Tipo de cuenta + Numero(crédito)
        '**************************************************
        loContenido.Append("000")
        loContenido.Append("0000000000000000000000")

        '**************************************************
        ' Encabezado fijo: código del proceso + relleno
        '**************************************************
        loContenido.Append("000000000000")
        loContenido.Append("000000000000000000000000000000000000000000000000")
        loContenido.AppendLine()


        '**************************************************
        ' Renglones
        '**************************************************
        Dim lnCantidad As Integer = loTabla.Rows.Count
        Dim lnTotalCredito As Decimal = 0D
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Datos: Número de renglon
            loContenido.Append(Strings.Right("000000" & CStr(n+1).Trim(), 6))

            ' Datos: Tipo de cuenta + Numero(crédito)
            loContenido.Append("CC ")
            loContenido.Append(Strings.Right("00000000000000000000" & CStr(loRenglon("Num_Cue")).Trim(), 20))

            'Datos: Cédula (17 caracteres, rellenar con espacios a la derecha)
            Dim lcCedula As String = CStr(loRenglon("Cedula")).ToUpper()
            lcCedula = loLimpiaCedula.Replace(lcCedula, "") & " "
            Dim lcNacionalidad As String = lcCedula(0)
            If (lcNacionalidad <> "V" AndAlso lcNacionalidad <> "E" )
                lcNacionalidad = "V"
            End If
            lcCedula = Strings.Right("0000000000" & loLimpiaNumero.Replace(lcCedula, ""), 10)
            loContenido.Append(lcNacionalidad)
            loContenido.Append(lcCedula)

            'Datos (FIJO): SERIAL-SERV + NUM-CUOTA + REFERENCIA
            loContenido.Append("00000")
            loContenido.Append("00000")
            loContenido.Append("0000000000")

            'Datos: Monto
            loContenido.Append(Strings.Right("000000000000000" & (CDec(loRenglon("Mon_Net")) * 100).ToString("0"), 15))
            
            lnTotalCredito += CDec(loRenglon("Mon_Net"))

            'Datos (FIJO): Tipo de Operación + Indicador de proceso
            loContenido.Append("C")
            loContenido.Append("0")

            'Datos: Descripción
            loContenido.Append(Strings.Left( CStr(loRenglon("Descripcion")) & Strings.Space(40), 40))

            'Datos (FIJO): Aplica cargo + Codigo Rechazo + Descripción Rechazo + Relleno
            loContenido.Append("0")
            loContenido.Append("000")
            loContenido.Append("0000000000000000000000000000000000000000")
            loContenido.Append("000000000")
            loContenido.AppendLine()
            
        Next

        '**************************************************
        ' Totales
        '**************************************************

        'Totales (FIJO): Consecutivo
        loContenido.Append("999999")

        'Totales: Nombre de la empresa
        loContenido.Append(Strings.Left( CStr(loRenglon("Nombre_Empresa")) & Strings.Space(40), 40))

        'Totales: cantidad de registros
        loContenido.Append(Strings.Right("000000" & CStr(lnCantidad).Trim(), 6))

        'Totales: total debitos
        loContenido.Append(Strings.Right("000000000000000" & (lnTotalCredito * 100).ToString("0"), 15))

        'Totales (FIJO): cantidad debitos
        loContenido.Append("000001")

        'Totales: total créditos
        loContenido.Append(Strings.Right("000000000000000" & (lnTotalCredito * 100).ToString("0"), 15))

        'Totales: cantidad créditos
        loContenido.Append(Strings.Right("000000" & CStr(lnCantidad).Trim(), 6))

        'Totales (FIJO): Relleno
        loContenido.AppendLine("0000000000000000000000000000000000000000000000000000000000000000000000000000")

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++




        ''**************************************************
        '' Encabezado variable 01: datos generales del TXT
        ''**************************************************
        'loContenido.Append("01")        'Tipo de renglon
        'loContenido.Append("SAL")       'FIJO

        'Dim lcCodigoDescripcionPago As String
        'lcCodigoDescripcionPago =  CStr(loRenglon("Codigo_Descripcion_Pago"))
        'lcCodigoDescripcionPago = loLimpiaNumero.Replace(lcCodigoDescripcionPago, "")
        'lcCodigoDescripcionPago =  Strings.Left(lcCodigoDescripcionPago & Strings.Space(2), 2)
        'loContenido.Append(lcCodigoDescripcionPago) 'Codigo de Descripcion del Pago

        'loContenido.Append("                              9  ")       'FIJO

        'Dim lcNumeroOrdenPago As String
        'lcNumeroOrdenPago =  CStr(loRenglon("Numero_Orden_Pago")).Trim().ToUpper()
        'lcNumeroOrdenPago = loLimpiaAlfaNum.Replace(lcNumeroOrdenPago, "")
        'lcNumeroOrdenPago =  Strings.Left(lcNumeroOrdenPago & Strings.Space(12), 12)
        'loContenido.Append(lcNumeroOrdenPago)           'Numero de la Orden de Pago (manual/opcional)

        'loContenido.Append("                       ")   'FIJO

        ''Encabezado: Fecha+hora actual (automático)
        'Dim lcFechaHora As String
        'Dim ldFechaHora As Date = Date.Now()
        'lcFechaHora = ldFechaHora.ToString("yyyyMMddHHmmss")
        'loContenido.Append(lcFechaHora)

        ''Encabezado: Fin encabezado 01
        'loContenido.AppendLine()

        ''**************************************************
        '' Encabezado variable 02: datos generales del TXT
        ''**************************************************
        'loContenido.Append("02")        'Tipo de renglon
        
        ''Encabezado: Referencia del débito (manual/opcional)
        'Dim lcReferenciaDelDebito As String
        'lcReferenciaDelDebito =  CStr(loRenglon("Referencia_Debito"))
        'lcReferenciaDelDebito = loLimpiaNumero.Replace(lcReferenciaDelDebito, "")
        'If (lcReferenciaDelDebito = "") Then lcReferenciaDelDebito = "5"
        'lcReferenciaDelDebito =  Strings.Right(Strings.StrDup(8, "0") & lcReferenciaDelDebito, 8)
        'loContenido.Append(lcReferenciaDelDebito)

        ''Encabezado: FIJO.
        'loContenido.Append("                      ")

        ''Encabezado: RIF del Ordenante (se toma de la empresa actual).
        'Dim lcRifOrdenante As String
        'lcRifOrdenante =  CStr(loRenglon("Rif_Empresa")).Trim()
        'lcRifOrdenante = loLimpiaRif.Replace(lcRifOrdenante, "")
        'lcRifOrdenante =  Strings.Left(lcRifOrdenante & Strings.Space(17), 17)
        'loContenido.Append(lcRifOrdenante)

        ''Encabezado: Nombre del Ordenante (se toma de la empresa actual). 
        'Dim lcNombreOrdenante As String
        'lcNombreOrdenante = CStr(loRenglon("Nombre_Empresa")).Trim()
        'lcNombreOrdenante = Me.mEliminarAcentos(lcNombreOrdenante)
        'lcNombreOrdenante = Strings.Left(lcNombreOrdenante & Strings.Space(35), 35)
        'loContenido.Append(lcNombreOrdenante)

        ''Encabezado: Monto Total (se obtiene de los renglones).
        'Dim lnMontoTotal As Long
        'For Each loItem As DataRow In loTabla.Rows
        '    lnMontoTotal += CLng(CDec(loItem("Mon_Net"))*100)
        'Next
        'Dim lcMontoTotal As String = lnMontoTotal.ToString("000000000000000")
        'loContenido.Append(lcMontoTotal)

        ''Encabezado: Moneda (fijo, 4 caracteres)
        'loContenido.Append("VEF ")

        ''Encabezado: Numero de cuenta de la empresa
        'Dim lcNumeroCuentaDebitar As String = CStr(loRenglon("Cuenta_Empresa")).Trim()
        'lcNumeroCuentaDebitar = loLimpiaNumero.Replace(lcNumeroCuentaDebitar, "")
        'lcNumeroCuentaDebitar = Strings.Left(lcNumeroCuentaDebitar & Strings.Space(20), 20)
        'loContenido.Append(lcNumeroCuentaDebitar)

        ''Encabezado: FIJO
        'loContenido.Append("              BANESCO    ")    

        ''Encabezado: Fecha de ejecución del pago
        'Dim lcFechaPago As String = ldFechaEmision.ToString("yyyyMMdd")
        'loContenido.Append(lcFechaPago)
        'loContenido.AppendLine()


        ''**************************************************
        '' Datos de trabajadores: montos a pagar
        ''**************************************************
        'Dim lnCantidad As Integer = loTabla.Rows.Count
        'For n As Integer = 0 To lnCantidad - 1
        '    loRenglon = loTabla.Rows(n)

        '    'Datos: Tipo de renglon (FIJO)
        '    loContenido.Append("03")

        '    'Datos: Referencia del crédito (30 caracteres, comienza formateado con 8 ceros, rellenar con espacios)
        '    Dim lcReferenciaCredito As String = (n+1).ToString("00000000")
        '    lcReferenciaCredito = Strings.Left(lcReferenciaCredito & Strings.Space(30), 30)
        '    loContenido.Append(lcReferenciaCredito)

        '    'Datos: Monto trabajador (15 caracteres, los dos últimos son decimales, rellenar con "0" a la izq.)
        '    Dim lnMonto As Long = CLng(CDec(loRenglon("Mon_Net"))*100)
        '    Dim lcMonto As String = lnMonto.ToString("000000000000000")
        '    loContenido.Append(lcMonto)

        '    'Datos: Moneda (fijo, 3 caracteres))
        '    loContenido.Append("VEF")

        '    'Datos: Cuenta (20 caracteres, rellenar con X en caso de error)
        '    Dim lcCuenta As String = CStr(loRenglon("Num_Cue")).Trim()
        '    lcCuenta = Strings.Left(lcCuenta & "XXXXXXXXXXXXXXXXXXXX", 20)
        '    loContenido.Append(lcCuenta)

        '    'Datos: FIJO (10 espacios)
        '    loContenido.Append("          ")

        '    'Datos: Código SUDEBAN del banco/Cuenta (primeros 4 caracteres de la cuenta, rellenar con X en caso de error)
        '    lcCuenta = Strings.Left(lcCuenta & "XXXX", 4)
        '    loContenido.Append(lcCuenta)

        '    'Datos: FIJO (10 espacios)
        '    loContenido.Append("          ")

        '    'Datos: Cédula (17 caracteres, rellenar con espacios a la derecha)
        '    Dim lcCedula As String = CStr(loRenglon("Cedula")).ToUpper()
        '    lcCedula = loLimpiaCedula.Replace(lcCedula, "")
        '    If (lcCedula = loLimpiaNumero.Replace(lcCedula, "") )
        '        lcCedula = "V" & lcCedula
        '    End If
        '    lcCedula = Strings.Left(lcCedula & Strings.Space(17), 17)
        '    loContenido.Append(lcCedula)

        '    'Datos: Nombre del trabajador (35 caracteres, rellenar con espacios)
        '    Dim lcNombre As String = CStr(loRenglon("Nom_Tra")).ToUpper()
        '    lcNombre = Me.mEliminarAcentos(lcNombre)
        '    lcNombre = Strings.Left(lcNombre & Strings.Space(35), 35)
        '    loContenido.Append(lcNombre)

        '    'Datos: FIJO (236 espacios)
        '    loContenido.Append(Strings.Space(236))

        '    'Datos: FIJO 
        '    loContenido.Append("42")

        '    'Datos: Cuenta Banesco (" ") u Otra ("5")
        '    If (lcCuenta = "0134") Then
        '        loContenido.Append(" ")
        '    Else
        '        loContenido.Append("5")
        '    End If
            
        '    'Fin de línea de detalle
        '    loContenido.Append(vbNewLine)
        
        'Next n
        
        ''Pie: Tipo de renglon
        'loContenido.Append("06")

        ''Pie: Cantidad de Débitos (15 dígitos, relleno con ceros a la izquierda)
        'loContenido.Append(1.ToString("000000000000000"))

        ''Pie: Cantidad de Créditos (15 dígitos, relleno con ceros a la izquierda)
        'loContenido.Append(loTabla.Rows.Count.ToString("000000000000000"))

        ''Pie: Monto Total (15 dígitos, relleno con ceros a la izquierda)
        ''loContenido.Append(lcMontoTotal)

        ''Pie: Fin de línea/archivo
        'loContenido.Append(vbNewLine)



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
        
        lcTexto = regex.Replace(lcTexto, "[^0-9a-zA-Z] ", "")

        Return lcTexto
    End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 27/07/15: Código Inicial.
'-------------------------------------------------------------------------------------------'
