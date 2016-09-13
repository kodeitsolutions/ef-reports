'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_AlimentacionTEBCA_SVT"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_AlimentacionTEBCA_SVT
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

        Dim lcOrden As String = goReportes.pcOrden

        'Parametros generación:
        Dim lnConsecutivo As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(2))
        lnConsecutivo = Math.Max(Math.Min(lnConsecutivo, 99), 1)
        Dim lcConsecutivo As String = lnConsecutivo.ToString("00")
        Dim lcConsecutivoSQL As String = goServicios.mObtenerCampoFormatoSQL(lcConsecutivo)

        Dim ldEmision As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcEmision As String = ldEmision.ToString("yyMMdd")
        Dim lcEmisionSQL As String = goServicios.mObtenerCampoFormatoSQL(ldEmision, goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)

        Dim lcNumeroLote As String = lcEmision & lcConsecutivo
        Dim lcNumeroLoteSQL As String = goServicios.mObtenerCampoFormatoSQL(lcNumeroLote)

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT      Recibos.Documento                                AS Documento,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                                  AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                             AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula                              AS Cedula,")
            loConsulta.AppendLine("            SUM(Renglones_Recibos.Mon_Net)                   AS Mon_Net,")
            loConsulta.AppendLine("            CAST(" & lcConsecutivoSQL & " AS CHAR(2))        AS Consecutivo,")
            loConsulta.AppendLine("            CAST(" & lcEmisionSQL & " AS DATE)               AS Emision,")
            loConsulta.AppendLine("            CAST(" & lcNumeroLoteSQL & " AS CHAR(8))         AS NumeroLote")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN Recibos ")
            loConsulta.AppendLine("        ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("        AND Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("    JOIN Trabajadores ")
            loConsulta.AppendLine("        ON Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("        AND Trabajadores.Status = 'A'")
            loConsulta.AppendLine("WHERE   Renglones_Recibos.Cod_Con IN ('A011', 'A016', 'A111','B011','B015')")
            loConsulta.AppendLine("    AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Recibos.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("GROUP BY Trabajadores.Nom_Tra,Trabajadores.Cod_Tra,Trabajadores.Cedula,Recibos.Documento ")
            loConsulta.AppendLine("ORDER BY Trabajadores.Cod_Tra ASC ")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_AlimentacionTEBCA_SVT", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_AlimentacionTEBCA_SVT.ReportSource = loObjetoReporte

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

    Private Sub mGenerarArchivoTxt(ByVal laDatosReporte As DataSet)
        Dim loTabla As DataTable = laDatosReporte.Tables(0)
        Dim loLimpiaCedula As New Regex("[^0-9]", RegexOptions.Compiled)

        If (loTabla.Rows.Count = 0) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.Rows(0)
        Dim loContenido As New StringBuilder()

        '**************************************************
        ' Primero el registro de control: ENCABEZADO
        '**************************************************
        'Encabezado: #Fijo "0"
        loContenido.Append("0")

        'Encabezado: Número de Lote (18 caracteres)
        Dim lcLote As String = CStr(loRenglon("NumeroLote"))
        loContenido.Append(lcLote)

        'Encabezado: Rif de la empresa (15 caracteres, rellenar con espacios)
        Dim lcRif As String = goEmpresa.pcRifEmpresa
        lcRif = Strings.Left(lcRif & Strings.Space(15), 15)
        loContenido.Append(lcRif)

        'Encabezado: Cantidad de Registros (5 caracteres, rellenar con ceros)
        Dim lnCantidad As Integer = loTabla.Rows.Count
        Dim lcCantidad As String = lnCantidad.ToString("00000")
        loContenido.Append(lcCantidad)

        'Encabezado: #Fijo "2" (Recarga)
        loContenido.Append("2")

        'Encabezado: Fecha emisión (8 caracteres)
        Dim lcFecha As String = CDate(loRenglon("Emision")).ToString("yyyyMMdd")
        loContenido.Append(lcFecha)

        'Encabezado: monto total (8 caracteres)
        Dim lnTotalMonto As Long = CLng(CDec(loTabla.Compute("SUM(Mon_Net)", "")) * 100)
        Dim lcTotalMonto As String = lnTotalMonto.ToString("000000000000000000")
        loContenido.Append(lcTotalMonto)

        'Fin de línea
        loContenido.Append(vbNewLine)

        '**************************************************
        ' Datos de trabajadores: MONTOS
        '**************************************************
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Datos: #Fijo "2"
            loContenido.Append("2")

            'Datos: Número de Lote (8 caracteres)
            loContenido.Append(lcLote)

            'Datos: Cédula
            Dim lcCedula As String = CStr(loRenglon("Cedula"))
            lcCedula = loLimpiaCedula.Replace(lcCedula, "")
            lcCedula = Strings.Left(lcCedula & Strings.Space(15), 15)
            loContenido.Append(lcCedula)

            'Datos: Monto trabajador
            Dim lnMonto As Long = CLng(CDec(loRenglon("Mon_Net")) * 100)
            Dim lcMonto As String = lnMonto.ToString("000000000000000000")
            loContenido.Append(lcMonto)

            If (n < lnCantidad - 1) Then
                'Fin de línea: excepto en el último registro
                loContenido.Append(vbNewLine)
            End If

        Next n


        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcLote & ".txt")
        Me.Response.ContentType = "text/plain"
        Me.Response.Write(loContenido.ToString())
        Me.Response.End()

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' EAG: 5/10/15: Código Inicial.
'-------------------------------------------------------------------------------------------'
