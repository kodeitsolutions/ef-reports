'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_AlimentacionTODOTICKET"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_AlimentacionTODOTICKET
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        
        Dim lcOrden As String = goReportes.pcOrden

        'Parametros generación:
        Dim lnConsecutivo As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(4))
        lnConsecutivo = Math.Max(Math.Min(lnConsecutivo, 99), 1)
        Dim lcConsecutivo As String = lnConsecutivo.ToString("00")
        Dim lcConsecutivoSQL As String = goServicios.mObtenerCampoFormatoSQL(lcConsecutivo)

        Dim ldEmision As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcEmisionSQL As String = goServicios.mObtenerCampoFormatoSQL(ldEmision, goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT      Recibos.Documento                                AS Documento,")
            loConsulta.AppendLine("            Recibos.Cod_Tra                                  AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                             AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula                              AS Cedula,")
            loConsulta.AppendLine("            Renglones_Recibos.Cod_Con                        AS Cod_Con,")
            loConsulta.AppendLine("            ROUND(Renglones_Recibos.Mon_Net, 2)              AS Mon_Net,")
            loConsulta.AppendLine("            Renglones_Recibos.Val_Car                        AS Val_Car,")
            loConsulta.AppendLine("            CAST(" & lcConsecutivoSQL & " AS CHAR(2))        AS Consecutivo,")
            loConsulta.AppendLine("            CAST(" & lcEmisionSQL & " AS DATE)               AS Emision,")
            loConsulta.AppendLine("            COALESCE(TodoTiket.Val_Car,'0000')               AS Codigo_Empresa")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN Recibos ")
            loConsulta.AppendLine("        ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("        AND Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("    JOIN Trabajadores ")
            loConsulta.AppendLine("        ON Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades TodoTiket")
            loConsulta.AppendLine("        ON  TodoTiket.Cod_Reg = " & goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo))
            loConsulta.AppendLine("        AND TodoTiket.Origen = 'Empresas'")
            loConsulta.AppendLine("        AND TodoTiket.Cod_Pro = 'CODTODOTKT'")
            loConsulta.AppendLine("WHERE   Renglones_Recibos.Cod_Con IN ('A011', 'A013', 'A311')")
            loConsulta.AppendLine("    AND Recibos.Fecha BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND Recibos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Trabajadores.Status IN (" & lcParametro3Desde & ")")
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_AlimentacionTODOTICKET", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_AlimentacionTODOTICKET.ReportSource = loObjetoReporte

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
        Dim loLimpiaCedula As New Regex("[^0-9]", RegexOptions.Compiled)

        If (loTabla.Rows.Count = 0 ) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.rows(0)
        Dim loContenido As New StringBuilder()

        'Fecha emisión (8 caracteres)
        Dim lcEmision As String = CDate(loRenglon("Emision")).ToString("ddMMyyyy")
        Dim lcConsecutivo As String = CStr(loRenglon("Consecutivo")).Trim()
        Dim lcCodigoEmpresa As String = CStr(loRenglon("Codigo_Empresa")).Trim()
        'Nombre de Archivo 
        Dim lcNombreArchivo As String = "ABONOS" & lcCodigoEmpresa & lcConsecutivo & lcEmision

        '**************************************************
        ' Datos de trabajadores:
        '**************************************************
        Dim lnCantidad As Integer = loTabla.Rows.Count 
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)
            
            'Datos: Tipo (V o E)
            Dim lcTipo As String = CStr(loRenglon("Cedula")).Trim()
            If (lcTipo.Length>0) Then
                lcTipo = Strings.Left(lcTipo, 1)
            End If
            If (lcTipo <> "V" AndAlso lcTipo <> "E")
                lcTipo = "X" 'Esto es para que produzca error al tratar de cargarlo: solo se permite V o E
            End If
            loContenido.Append(lcTipo)

            'Datos: Cédula (9 dígitos, rellenar con 0 a la izquierda)
            'Si hay más de 9 se envía todo (al subir el TXT se validará) 
            Dim lcCedula As String = CStr(loRenglon("Cedula")).Trim()
            lcCedula = loLimpiaCedula.Replace(lcCedula, "")
            If (lcCedula.Length < 9) Then
                lcCedula = Strings.Right("000000000" & lcCedula, 9)
            End If
            loContenido.Append(lcCedula)

            'Datos: "Relleno" (2 espacios)
            loContenido.Append("  ")

            'Datos: Monto trabajador (21 dígitos sin separadores, los dos últimos son decimales)
            Dim lnMonto As Long = CLng(CDec(loRenglon("Mon_Net"))*100)
            Dim lcMonto As String = lnMonto.ToString("000000000000000000000")
            loContenido.Append(lcMonto)

            'Datos: Fecha de recarga (8 caracteres, DDMMYYYY)
            loContenido.Append(lcEmision)
            
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

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 17/09/14: Código Inicial.
'-------------------------------------------------------------------------------------------'
