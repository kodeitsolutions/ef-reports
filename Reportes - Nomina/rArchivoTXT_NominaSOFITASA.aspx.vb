'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_NominaSOFITASA"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_NominaSOFITASA
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
        
        Dim lcRifEmpresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcRifEmpresa)

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cedula                             AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                            AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                            AS Nom_Tra,")
            loConsulta.AppendLine("            COALESCE(Prop_Primer_Nombre.Val_Car, '')        AS Primer_Nombre,")
            loConsulta.AppendLine("            COALESCE(Prop_Primer_Apellido.Val_Car, '')      AS Primer_Apellido,")
            loConsulta.AppendLine("            Trabajadores.Num_Cue                            AS Num_Cue,")
            loConsulta.AppendLine("            ROUND(Pagos.Mon_Net, 2)                         AS Mon_Net,")
            loConsulta.AppendLine("            CAST( " & lcParametro3Desde & " AS DATE)        AS Emision,")
            loConsulta.AppendLine("            CAST( " & lcParametro4Desde & " AS INT)         AS Correlativo,")
            loConsulta.AppendLine("            CAST( " & lcRifEmpresa & " AS VARCHAR(20))      AS Rif_Empresa")
            loConsulta.AppendLine("FROM        Trabajadores")
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
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Primer_Nombre")
            loConsulta.AppendLine("        ON  Prop_Primer_Nombre.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Cod_Pro = 'NOMTRA01'")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Primer_Apellido")
            loConsulta.AppendLine("        ON  Prop_Primer_Apellido.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Cod_Pro = 'NOMTRA03'")
            loConsulta.AppendLine("WHERE   Pagos.Mon_Net > 0")
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_NominaSOFITASA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_NominaSOFITASA.ReportSource = loObjetoReporte

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

        If (loTabla.Rows.Count = 0 ) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.rows(0)
        Dim ldFechaEmision As Date = CDate(loRenglon("Emision"))
        Dim lcFechaEmision As String = ldFechaEmision.ToString("ddMMyyyy")
        Dim lcNombreArchivo As String = "NOMINA_SOFITASA_" & ldFechaEmision.ToString("ddMMyy")

        Dim loContenido As New StringBuilder()

        '**************************************************
        ' Datos de la empresa: cabecera resumen
        '**************************************************

        'Cabecera: Tipo (valor fijo "0")
        loContenido.Append("0")

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

        lcRif = Strings.Right("000000000000000" & lcRif, 15)

        loContenido.Append(lcTipoRif)
        loContenido.Append(lcRif)

        'Cabecera: Fecha de Generación del pago (8 caracteres, formato ddMMyyyy)
        loContenido.Append(lcFechaEmision)

        'Cabecera: Correlativo (10 dígitos, rellenar con cero a la izquierda)
        Dim lnCorrelativo AS Integer = CInt(loRenglon("Correlativo"))
        loContenido.Append(lnCorrelativo.ToString("0000000000"))

        'Cabecera: N° de registros (5 dígitos, rellenar con cero a la izquierda)
        Dim lnCantidad As Integer = loTabla.Rows.Count
        loContenido.Append(lnCantidad.ToString("00000"))

        'Cabecera: Monto Total (15 dígitos, los dos últimos son decimales, rellenar con cero a la izquierda)
        Dim lnMontoTotal As Long = CLng(CDec(loTabla.Compute("SUM(Mon_Net)", ""))*100)
        loContenido.Append(lnMontoTotal.ToString("000000000000000"))

        'Cabecera: Fin de línea
        loContenido.Append(vbNewLine)

        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Datos: Tipo (valor fijo "1")
            loContenido.Append("1")

            'Datos: Cédula (15 caracteres, rellenar con 0 entre el tipo y el número)
            Dim lcCedula As String = CStr(loRenglon("Cedula")).ToUpper()
            Dim lcTipoCI As String 
            If (lcCedula.Length = 0) Then 
                lcTipoCI = "V"
                lcCedula = ""
            Else
                lcTipoCI = lcCedula.Substring(0, 1)
                lcCedula = loSoloNumeros.Replace(lcCedula, "")
            End If

            lcCedula = Strings.Right("000000000000000" & lcCedula, 15)

            loContenido.Append(lcTipoCI)
            loContenido.Append(lcCedula)


            'Datos: Cuenta (20 caracteres, rellenar con X en caso de error)
            Dim lcCuenta As String = CStr(loRenglon("Num_Cue")).Trim()
            lcCuenta = loSoloNumeros.Replace(lcCuenta, "")
            lcCuenta = Strings.Left(lcCuenta & "XXXXXXXXXXXXXXXXXXXX", 20)
            loContenido.Append(lcCuenta)

            'Datos: Monto trabajador (15 dígitos, los dos últimos son decimales, rellenar con "0" a la izq.)
            Dim lnMonto As Long = CLng(CDec(loRenglon("Mon_Net"))*100)
            Dim lcMonto As String = lnMonto.ToString("000000000000000")
            loContenido.Append(lcMonto)

            'Datos: Nombre del trabajador (Primer Apellido y Primer Nombre)
            Dim lcApellido As String = CStr(loRenglon("Primer_Apellido")).ToUpper().Trim()
            Dim lcNombre As String = CStr(loRenglon("Primer_Nombre")).ToUpper().Trim()
            loContenido.Append(lcApellido)
            loContenido.Append(" ")
            loContenido.Append(lcNombre)
            
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
' RJG: 18/09/14: Código Inicial.
'-------------------------------------------------------------------------------------------'
