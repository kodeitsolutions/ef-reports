'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MHL_Txt_Val_Equis"
'-------------------------------------------------------------------------------------------'
Partial Class MHL_Txt_Val_Equis

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpArticulos( RTG         VARCHAR(30),")
            loConsulta.AppendLine("                            Tipo_Mer    CHAR(2),")
            loConsulta.AppendLine("                            Articulo    VARCHAR(30),")
            loConsulta.AppendLine("                            Unidades    DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Bruto       DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Neto        DECIMAL(28, 10));")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Ventas:")
            loConsulta.AppendLine("INSERT INTO #tmpArticulos(RTG, Tipo_Mer, Articulo, Unidades, Bruto, Neto)")
            loConsulta.AppendLine("SELECT      Clientes.Atributo_D                                             AS RTG, ")
            loConsulta.AppendLine("            CASE")
            loConsulta.AppendLine("                WHEN Clientes.Atributo_D = 'RTG' THEN '20' ")
            loConsulta.AppendLine("                ELSE '10' ")
            loConsulta.AppendLine("            END                                                             AS Tipo_Mer,")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Art                                      AS Articulo, ")
            loConsulta.AppendLine("            Renglones_Facturas.Can_Art1                                     AS Unidades, ")
            loConsulta.AppendLine("            (CASE WHEN renglones_facturas.Usa_Des_Com = 1")
            loConsulta.AppendLine("                THEN ( '0' + RTRIM((CONVERT(xml, renglones_facturas.Des_Com)).query('//descuentos/parametros/pre_ori').value('.', 'VARCHAR(MAX)')) )")
            loConsulta.AppendLine("                ELSE (  CASE WHEN renglones_facturas.Usa_Des_Vol = 1")
            loConsulta.AppendLine("                            THEN ( '0' + RTRIM((CONVERT(xml, renglones_facturas.Des_Vol)).query('//descuentos/parametros/pre_ori').value('.', 'VARCHAR(MAX)')) )")
            loConsulta.AppendLine("                            ELSE renglones_facturas.precio1")
            loConsulta.AppendLine("                        END)")
            loConsulta.AppendLine("            END)*Renglones_Facturas.Can_Art1                                AS Bruto,")
            loConsulta.AppendLine("            CAST(Renglones_Facturas.mon_net AS DECIMAL(28,10))              AS Neto")
            loConsulta.AppendLine("FROM        Renglones_Facturas ")
            loConsulta.AppendLine("    JOIN    Facturas    ON  Facturas.Documento  =   Renglones_Facturas.Documento ")
            loConsulta.AppendLine("    JOIN    Clientes    ON  Clientes.Cod_Cli    =   Facturas.Cod_Cli ")
            loConsulta.AppendLine("    JOIN    Articulos   ON  Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art ")
            loConsulta.AppendLine("WHERE       Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 		       AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 		AND Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine(" 		AND Articulos.Cod_dep NOT IN ('XGAS', 'XIMP', 'ZCOM');")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Devoluciones:")
            loConsulta.AppendLine("INSERT INTO #tmpArticulos(RTG, Tipo_Mer, Articulo, Unidades, Bruto, Neto)")
            loConsulta.AppendLine("SELECT	    Clientes.Atributo_D                                             AS RTG, ")
            loConsulta.AppendLine("            CASE")
            loConsulta.AppendLine("                WHEN Clientes.Atributo_D = 'RTG' THEN '20' ")
            loConsulta.AppendLine("                ELSE '10' ")
            loConsulta.AppendLine("            END                                                             AS Tipo_Mer,")
            loConsulta.AppendLine("            Renglones_dClientes.Cod_Art                                     AS Articulo, ")
            loConsulta.AppendLine("            -Renglones_dClientes.Can_Art1                                   AS Unidades, ")
            loConsulta.AppendLine("            CAST(Renglones_dClientes.precio1 AS DECIMAL(28,10))")
            loConsulta.AppendLine("                *Renglones_dClientes.Can_Art1                               AS Bruto,")
            loConsulta.AppendLine("            CAST(Renglones_dClientes.Mon_Net AS DECIMAL(28,10))             AS Neto")
            loConsulta.AppendLine("FROM        Renglones_dClientes ")
            loConsulta.AppendLine("   JOIN     Devoluciones_clientes on Devoluciones_clientes.Documento  =   Renglones_dClientes.documento ")
            loConsulta.AppendLine("   left JOIN     Clientes    ON  Clientes.Cod_Cli    =   Devoluciones_Clientes.Cod_Cli ")
            loConsulta.AppendLine("   left JOIN     Articulos   ON  Articulos.Cod_Art   =   Renglones_dClientes.Cod_Art ")
            loConsulta.AppendLine("WHERE       Devoluciones_clientes.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 		       AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 		AND Devoluciones_Clientes.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine(" 		AND Articulos.Cod_dep NOT IN ('XGAS','XIMP','ZCOM');")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Datos de salida:")
            loConsulta.AppendLine("SELECT      RTG                                     AS RTG,")
            loConsulta.AppendLine("            Tipo_Mer                                AS Tipo_Mer,")
            loConsulta.AppendLine("            Articulo                                AS Articulo,")
            loConsulta.AppendLine("            CAST(SUM(Unidades) AS DECIMAL(28,0))    AS Unidades,")
            loConsulta.AppendLine("            CAST(SUM(Bruto) AS DECIMAL(28,0))       AS Mon_Bru,")
            loConsulta.AppendLine("            CAST(SUM(Neto) AS DECIMAL(28,0))        AS Mon_Net")
            loConsulta.AppendLine("FROM        #tmpArticulos")
            loConsulta.AppendLine("GROUP BY    Articulo, RTG, Tipo_Mer")
            loConsulta.AppendLine("ORDER BY    Articulo, RTG;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpArticulos;")
            loConsulta.AppendLine("")

            Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------'
            ' A partir de aqui se genera el archivo de texto				                                        ' 
            '-------------------------------------------------------------------------------------------------------'
            Dim ldFechaFin AS Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcPeriodo As String = Strings.Format(ldFechaFin, "yyyyMM")
            Dim loSalida As New StringBuilder()
            Dim lnContador As Integer = 0
            
            For Each loFila As DataRow In laDatosReporte.Tables(0).Rows
                
                'Encabezado:
                If (lnContador = 0) Then
                    Dim ldFechaExpedicion As Date = New Date()

                    Dim lcAnio As String = ldFechaFin.ToString("yyyy")
                    Dim lcMes As String = ldFechaFin.ToString("MM")
                    Dim lcAnio_Exp As String = ldFechaExpedicion.ToString("yyyy")
                    Dim lcMes_Exp As String = ldFechaExpedicion.ToString("MM")
                    Dim lcDia_Exp As String = ldFechaExpedicion.ToString("dd")

                    loSalida.Append("1VENTES")
                    loSalida.Append(lcAnio)
                    loSalida.Append(lcMes)
                    loSalida.Append("VE  VAL2")
                    loSalida.Append(lcAnio_Exp)
                    loSalida.Append(lcMes_Exp)
                    loSalida.Append(lcDia_Exp)
                    loSalida.Append("VEF")
                    loSalida.Append("        ")
                    loSalida.Append("                    ")
                    loSalida.Append("                    ")
                    loSalida.Append("                    ")
                    loSalida.Append("                    ")
                    loSalida.AppendLine()

                End If
                
                'Detalle
                Dim lcArticulo As String = Strings.Right("      " + CStr(loFila("Articulo")).Trim(), 6) 'Normalmente 6c
                Dim lcTipo_Mer As String = Strings.Right("  " + CStr(loFila("Tipo_Mer")).Trim(), 2)     '3c
                Dim lcRTG As String = Strings.Right("   " + CStr(loFila("RTG")).Trim(), 3)              '2c
                Dim lcUnidades As String = CDec(loFila("Unidades")).ToString("0000000")                 '7c
                Dim lcMontoNeto As String = CDec(loFila("Mon_Net")).ToString("0000000000000")           '13c
                Dim lcMontoBruto As String = CDec(loFila("Mon_Bru")).ToString("0000000000000")          '13c
                Dim lcMontoNAR As String = lcMontoBruto                                                 '13c
                Dim lcMontoCRV As String = lcMontoBruto                                                 '13c

                loSalida.Append("2")
                loSalida.Append(lcArticulo)
                loSalida.Append("484")
                loSalida.Append(lcTipo_Mer)
                loSalida.Append(lcRTG)
                loSalida.Append("  +")
                loSalida.Append(lcUnidades)
                loSalida.Append("  +")
                loSalida.Append(lcMontoNeto)
                loSalida.Append("  +")
                loSalida.Append(lcMontoBruto)
                loSalida.Append("  +")
                loSalida.Append(lcMontoNAR)
                loSalida.Append("  +")
                loSalida.Append("0000000000000")
                loSalida.Append("  +")
                loSalida.Append(lcMontoNAR)
                loSalida.Append("VEF                      ")

                loSalida.AppendLine()
                lnContador = lnContador + 1
            Next loFila

            loSalida.Append( (90000 + lnContador).ToString("00000"))

            '-------------------------------------------------------------------------------------------------------
            ' Envia la salida a pantalla a un archivo descargable.
            '-------------------------------------------------------------------------------------------------------
            Me.Response.Clear()
            Me.Response.ContentEncoding = System.Text.Encoding.UTF8
            Me.Response.AppendHeader("content-disposition", "attachment; filename=F336VEN" & ".txt")
            Me.Response.ContentType = "plain/text"
            Me.Response.Write(loSalida.ToString())
            'Me.Response.Write(Strings.Space(20))	'A veces no todo el texto es enviado a pantalla, entonces se 
            Me.Response.End()                       'mandan algunos espacios en blanco adicionales para "rellenar".

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/08/09: Codigo inicial (creado a partir de rCRetencion_IVAProveedores)				'
'-------------------------------------------------------------------------------------------'
' JJD: 08/06/11: Ajustes en la generacion de monto base y porcentaje.				        '
'-------------------------------------------------------------------------------------------'
' MAT: 10/06/11: Generación del archivo del comprobante con cero retención				    '
'-------------------------------------------------------------------------------------------'
' JJD: 04/12/13: Ajuste para la generacion del reporte VAL X				                '
'-------------------------------------------------------------------------------------------'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               