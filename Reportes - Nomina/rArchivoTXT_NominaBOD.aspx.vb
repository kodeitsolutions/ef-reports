'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_NominaBOD"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_NominaBOD
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        
        Dim lcOrden As String = goReportes.pcOrden
        
        Dim lcRifEmpresaSQL As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)
        Dim lcCodigoEmpresaSQL As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cedula                                 AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                                AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                                AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Num_Cue                                AS Num_Cue,")
            loConsulta.AppendLine("            Trabajadores.Correo                                 AS Email_Trabajador,")
            loConsulta.AppendLine("            Trabajadores.Movil                                  AS Movil_Trabajador,")
            loConsulta.AppendLine("            ROUND(Pagos.Mon_Net, 2)                             AS Mon_Net,")
            loConsulta.AppendLine("            Pagos.Documento                                     AS Documento,")
            loConsulta.AppendLine("            Pagos.Fec_Ini                                       AS Fecha_Pago,")
            loConsulta.AppendLine("            Pagos.Comentario                                    AS Comentario,")
            loConsulta.AppendLine("            CAST( " & lcParametro3Desde & " AS DATE)            AS Emision,")
            loConsulta.AppendLine("            CAST( " & lcParametro4Desde & " AS INT)             AS Numero_Lote,")
            loConsulta.AppendLine("            COALESCE(Prop_Numero_Contrato.Val_Car, '')          AS Numero_Contrato,")
            loConsulta.AppendLine("            CAST( " & lcRifEmpresaSQL & " AS VARCHAR(20))       AS Rif_Empresa")
            loConsulta.AppendLine("FROM        Trabajadores")
            loConsulta.AppendLine("    JOIN  ( SELECT  SUM(Recibos.Mon_Net) AS Mon_Net,")
            loConsulta.AppendLine("                    Recibos.Cod_Tra,")
            loConsulta.AppendLine("                    Recibos.Documento,")
            loConsulta.AppendLine("                    Recibos.Fec_Ini,")
            loConsulta.AppendLine("                    Recibos.Comentario")
            loConsulta.AppendLine("            FROM    Recibos")
            loConsulta.AppendLine("            WHERE   Recibos.Cod_Con NOT IN  ('92','93','94','95')")
            loConsulta.AppendLine("                AND Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("                AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                AND Recibos.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                AND " & lcParametro1Hasta)
            loConsulta.AppendLine("            GROUP BY Recibos.Cod_Tra, Recibos.Documento, Recibos.Fec_Ini, Recibos.Comentario")
            loConsulta.AppendLine("            ) AS Pagos")
            loConsulta.AppendLine("        ON  Pagos.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Numero_Contrato")
            loConsulta.AppendLine("        ON  Prop_Numero_Contrato.Cod_Reg = " & lcCodigoEmpresaSQL)
            loConsulta.AppendLine("        AND Prop_Numero_Contrato.Origen = 'Empresas'")
            loConsulta.AppendLine("        AND Prop_Numero_Contrato.Cod_Pro = 'NUMCONBOD'")
            loConsulta.AppendLine("WHERE   Pagos.Mon_Net > 0")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Ban = 'BOD' ")
            loConsulta.AppendLine("    AND Trabajadores.Tip_Pag = 'Transferencia' ")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Tra BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrden)
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_NominaBOD", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_NominaBOD.ReportSource = loObjetoReporte

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
        Dim loSoloNumeros As New Regex("[^0-9]", RegexOptions.Compiled)
        Dim loAlfaNumerico As New Regex("[^/a-zA-Z. ]", RegexOptions.Compiled)

        If (loTabla.Rows.Count = 0 ) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.rows(0)
        Dim ldFechaEmision As Date = CDate(loRenglon("Emision"))
        'Dim lcFechaEmision As String = ldFechaEmision.ToString("ddMMyyyy")
        Dim lcNombreArchivo As String = "NOMINA_BOD_" & ldFechaEmision.ToString("ddMMyy")

        Dim loContenido As New StringBuilder()

        '**************************************************
        ' Datos de la empresa: cabecera resumen
        '**************************************************

        'Cabecera: Tipo de Registro (valor fijo "01")
        loContenido.Append("01")

        'Cabecera: Descripción del Lote (valor fijo "NOMINA") (20 caracteres, rellenar con espacios a la der.)
        Dim lcDescripcionLote As String = "NOMINA"
        lcDescripcionLote = Strings.Left(lcDescripcionLote & Strings.Space(20), 20)
        loContenido.Append(lcDescripcionLote)

        'Cabecera: RIF Empresa (15 caracteres, rellenar con 0 entre el tipo y el número)
        Dim lcRif As String = CStr(loRenglon("Rif_Empresa")).Trim()
        Dim lcTipoRif As String 
        If (lcRif.Length = 0) Then 
            lcTipoRif = "J"
            lcRif = ""
        Else
            lcTipoRif = lcRif.Substring(0, 1)
            lcRif = loSoloNumeros.Replace(lcRif, "")
        End If

        lcRif = Strings.Right("000000000" & lcRif, 9)

        loContenido.Append(lcTipoRif)
        loContenido.Append(lcRif)

        'Cabecera: Numero de Contrato (Asignado por BOD) (17 dígitos, rellenar con ceros a la izq.)
        Dim lcNumeroContrato As String = CStr(loRenglon("Numero_Contrato")).Trim()
        lcNumeroContrato = loSoloNumeros.Replace(lcNumeroContrato, "")
        lcNumeroContrato = Strings.Right(Strings.StrDup(17, "0") & lcNumeroContrato, 17)
        loContenido.Append(lcNumeroContrato)

        'Cabecera: Numero de Lote (manual, asignado por usuario) (9 dígitos, rellenar con ceros a la izq.)
        Dim lcNumeroLote As String = CStr(loRenglon("Numero_Lote")).Trim()
        lcNumeroLote = loSoloNumeros.Replace(lcNumeroLote, "")
        lcNumeroLote = Strings.Right(Strings.StrDup(9, "0") & lcNumeroLote, 9)
        loContenido.Append(lcNumeroLote)

        'Cabecera: Fecha de Envío (8 caracteres, formato ddMMyyyy)
        Dim lcFechaEnvio As String = ldFechaEmision.ToString("yyyyMMdd")
        loContenido.Append(lcFechaEnvio)

        'Cabecera: N° de registros (6 dígitos, rellenar con ceros a la izq)
        Dim lnCantidad As Integer = loTabla.Rows.Count
        Dim lcCantidad As String = lnCantidad.ToString(Strings.StrDup(6, "0"))
        loContenido.Append( lcCantidad )

        'Cabecera: Monto Total (17 dígitos, los dos últimos son decimales, rellenar con cero a la izquierda)
        Dim lnMontoTotal As Long = CLng(CDec(loTabla.Compute("SUM(Mon_Net)", ""))*100)
        loContenido.Append(lnMontoTotal.ToString(Strings.StrDup(17, "0")))

        'Cabecera: Moneda (3 caracteres, valor fijo "VEB")
        loContenido.Append(lnMontoTotal.ToString("VEB"))

        'Cabecera: Relleno (158 espacios)
        loContenido.Append(Strings.StrDup(158, " "))

        'Cabecera: Fin de línea
        loContenido.Append(vbNewLine)

        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Datos: Tipo de Registro (valor fijo "02")
            loContenido.Append("02")

            'Datos: Cédula (9 caracteres, rellenar con 0 entre el tipo y el número)
            Dim lcCedula As String = CStr(loRenglon("Cedula")).ToUpper()
            Dim lcTipoCI As String 
            If (lcCedula.Length = 0) Then 
                lcTipoCI = "V"
                lcCedula = ""
            Else
                lcTipoCI = lcCedula.Substring(0, 1)
                lcCedula = loSoloNumeros.Replace(lcCedula, "")
            End If

            lcCedula = Strings.Right(Strings.StrDup(9, "0") & lcCedula, 9)

            loContenido.Append(lcTipoCI)
            loContenido.Append(lcCedula)

            'Datos: Nombre del trabajador (60 caracteres, rellenar con espacios)
            Dim lcNombre As String = CStr(loRenglon("Nom_Tra")).ToUpper().Trim()
            lcNombre = Me.mConvertirANSI(lcNombre)
            lcNombre = loAlfaNumerico.Replace(lcNombre, "")
            lcNombre = Strings.Left(lcNombre & Strings.StrDup(60, " "), 60)
            loContenido.Append(lcNombre)

            'Datos: Referencia de la operacion (9 dígitos, rellenar con 0 a la izq)
            Dim lcReferenciaOperacion As String = CStr(loRenglon("Documento")).Trim()
            lcReferenciaOperacion = loSoloNumeros.Replace(lcReferenciaOperacion, "")
            lcReferenciaOperacion = Strings.Right(strings.StrDup(9, "0") & lcReferenciaoperacion, 9)
            loContenido.Append(lcReferenciaOperacion)

            'Datos: Descripción de la Operación (30 caracteres, rellenar con espacios)
            Dim lcDescripcionOperacion As String = CStr(loRenglon("Comentario")).Trim()
            lcDescripcionOperacion = Me.mConvertirANSI(lcDescripcionOperacion)
            lcDescripcionOperacion = loAlfaNumerico.Replace(lcDescripcionOperacion, "")
            lcDescripcionOperacion = Strings.Left(lcDescripcionOperacion & Strings.StrDup(30, " "), 30)
            loContenido.Append(lcDescripcionOperacion)

            'Datos: Modalidad de Pago (3 caracteres: CTA, BAN, EFE, TAR, AME, V/M)
            loContenido.Append("CTA")

            'Datos: Número de Cuenta (20 caracteres, rellenar con 0)
            Dim lcCuenta As String = CStr(loRenglon("Num_Cue")).Trim()
            lcCuenta = loSoloNumeros.Replace(lcCuenta, "")
            lcCuenta = Strings.Right(Strings.StrDup(20, "0") & lcCuenta, 20)
            loContenido.Append(lcCuenta)

            'Datos: Código del banco (4 dígitos (¿fijo?): "0137" -> BOD)
            loContenido.Append("0137")

            'Datos: Fecha Valor (8 caracteres, formato yyyyMMdd)
            Dim lcFechaValor As String = CDate(loRenglon("Fecha_Pago")).ToString("yyyyMMdd")
            loContenido.Append(lcFechaValor)

            'Datos: Monto trabajador (15 dígitos, los dos últimos son decimales, rellenar con "0" a la izq.)
            Dim lnMonto As Long = CLng(CDec(loRenglon("Mon_Net"))*100)
            Dim lcMonto As String = lnMonto.ToString(Strings.StrDup(15, "0"))
            loContenido.Append(lcMonto)
                
            'Datos: Moneda (3 caracteres, valor fijo "VEB")
            loContenido.Append(lnMontoTotal.ToString("VEB"))

            'Datos: impuesto retenido (15 dígitos, los dos últimos son decimales, rellenar con "0" a la izq.)
            lcMonto = Strings.StrDup(15, "0")
            loContenido.Append(lcMonto)

            'Datos: Descripción de la Operación (40 caracteres, rellenar con espacios)
            Dim lcEmail As String = CStr(loRenglon("Email_Trabajador")).Trim()
            lcEmail = Me.mConvertirANSI(lcEmail)
            lcEmail = Strings.Left(lcEmail & Strings.StrDup(40, " "), 40)
            loContenido.Append(lcEmail)

            'Datos: Descripción de la Operación (40 caracteres, rellenar con espacios)
            Dim lcCelular As String = CStr(loRenglon("Movil_Trabajador")).Trim()
            lcCelular = loSoloNumeros.Replace(lcCelular, "")
            lcCelular = Strings.Left(lcCelular & Strings.StrDup(11, " "), 11)
            loContenido.Append(lcCelular)

            'Datos: Relleno (20 espacios)
            loContenido.Append(Strings.StrDup(20, " "))

            If (n < lnCantidad-1) Then
                'Fin de línea: excepto en el último registro
                loContenido.Append(vbNewLine)
            End if
    
        Next n
        
        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcNombreArchivo & ".txt")
        Me.Response.ContentType = "text/plain"
        Me.Response.Write(loContenido.ToString())
        Me.Response.End()

    End Sub

    ''' <summary>
    ''' Sustituye algunos caracteres comúnes no válidos en ASCII/ANSI por
    ''' su equivalente ASCII.
    ''' </summary>
    ''' <param name="lcCadena">Cadena que se va a convertir</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function mConvertirANSI(lcCadena As String) As String

        lcCadena = lcCadena.Replace("á", "a").Replace("Á", "A")
        lcCadena = lcCadena.Replace("ä", "a").Replace("Ä", "A")
        lcCadena = lcCadena.Replace("à", "a").Replace("À", "A")
        lcCadena = lcCadena.Replace("é", "e").Replace("É", "E")
        lcCadena = lcCadena.Replace("ë", "e").Replace("Ë", "E")
        lcCadena = lcCadena.Replace("è", "e").Replace("È", "E")
        lcCadena = lcCadena.Replace("í", "i").Replace("Í", "I")
        lcCadena = lcCadena.Replace("ï", "i").Replace("Ï", "I")
        lcCadena = lcCadena.Replace("ì", "i").Replace("Ì", "I")
        lcCadena = lcCadena.Replace("ó", "o").Replace("Ó", "O")
        lcCadena = lcCadena.Replace("ö", "o").Replace("Ö", "O")
        lcCadena = lcCadena.Replace("ò", "o").Replace("Ò", "O")
        lcCadena = lcCadena.Replace("ú", "u").Replace("Ú", "U")
        lcCadena = lcCadena.Replace("ü", "u").Replace("Ü", "U")
        lcCadena = lcCadena.Replace("ù", "u").Replace("Ù", "U")
        lcCadena = lcCadena.Replace("ñ", "n").Replace("Ñ", "N")
        lcCadena = lcCadena.Replace("ç", "c").Replace("Ç", "C")

        Return lcCadena 
    End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 18/09/14: Código Inicial.
'-------------------------------------------------------------------------------------------'
