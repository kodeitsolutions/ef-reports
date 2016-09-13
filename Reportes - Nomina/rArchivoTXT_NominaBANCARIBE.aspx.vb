'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_NominaBANCARIBE"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_NominaBANCARIBE
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

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
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

        Dim lcOrden As String = goReportes.pcOrden

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cedula                     AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                    AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                    AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Num_Cue                    AS Num_Cue,")
            loConsulta.AppendLine("            ROUND(Pagos.Mon_Net, 2)                 AS Mon_Net")
            loConsulta.AppendLine("FROM        Trabajadores")
            loConsulta.AppendLine("    JOIN  ( SELECT  SUM(Recibos.Mon_Net) AS Mon_Net,")
            loConsulta.AppendLine("                    Recibos.Cod_Tra")
            loConsulta.AppendLine("            FROM    Recibos")
            loConsulta.AppendLine("            WHERE   Recibos.Cod_Con NOT IN  ('92','93','94','95')")
            loConsulta.AppendLine("                AND Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("                AND Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("                AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("                AND Recibos.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                AND " & lcParametro1Hasta)
            loConsulta.AppendLine("            GROUP BY Recibos.Cod_Tra")
            loConsulta.AppendLine("            ) AS Pagos")
            loConsulta.AppendLine("        ON  Pagos.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("WHERE   Pagos.Mon_Net > 0")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Tra BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Trabajadores.Cod_Con BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("    AND " & lcParametro4Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrden)
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_NominaBANCARIBE", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_NominaBANCARIBE.ReportSource = loObjetoReporte

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
        Dim loSoloNumeros As New Regex("[^0-9]", RegexOptions.Compiled)
        Dim loLimpiaCedula As New Regex("[^VE0-9]", RegexOptions.Compiled)

        If (loTabla.Rows.Count = 0) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.Rows(0)
        Dim lcNombreArchivo As String = "NOMINA_BANCARIBE_" & Date.Today.ToString("dd_MM_yy")

        Dim loContenido As New StringBuilder()

        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        Dim lnCantidad As Integer = loTabla.Rows.Count
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)


            'Datos: Cuenta (20 caracteres, rellenar con X en caso de error)
            Dim lcCuenta As String = CStr(loRenglon("Num_Cue")).Trim()
            lcCuenta = loSoloNumeros.Replace(lcCuenta, "")
            lcCuenta = Strings.Right(Strings.StrDup(20, "0") & lcCuenta, 20) & "/"
            loContenido.Append(lcCuenta)

            'Datos: Monto trabajador (15 caracteres, los dos últimos son decimales, rellenar con "0" a la izq.)

            Dim lcMonto As String = CDec(loRenglon("Mon_Net")).ToString("00,000.00") & "/"
            loContenido.Append(lcMonto)

            'Datos: Nombre del trabajador (35 caracteres, rellenar con espacios)
            Dim lcNombre As String = CStr(loRenglon("Nom_Tra")).ToUpper()
            lcNombre = Strings.Left(lcNombre & Strings.Space(35), 35)
            loContenido.Append(lcNombre)


            If (n < lnCantidad - 1) Then
                'Fin de línea: excepto en el último registro
                loContenido.Append(vbNewLine)
            End If

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
' EAG: 09/09/15: Código Inicial.
'-------------------------------------------------------------------------------------------'
