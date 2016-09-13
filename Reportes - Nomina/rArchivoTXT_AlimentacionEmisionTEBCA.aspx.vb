'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoTXT_AlimentacionEmisionTEBCA"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoTXT_AlimentacionEmisionTEBCA
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        
        Dim lcOrden As String = goReportes.pcOrden

        'Parametros generación:
        Dim lnConsecutivo As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(2))
        lnConsecutivo = Math.Max(Math.Min(lnConsecutivo, 99), 1)
        Dim lcConsecutivo As String = lnConsecutivo.ToString("00")
        Dim lcConsecutivoSQL As String = goServicios.mObtenerCampoFormatoSQL(lcConsecutivo)

        Dim ldEmision As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcEmision As String = ldEmision.ToString("yyMMdd")
        Dim lcEmisionSQL As String = goServicios.mObtenerCampoFormatoSQL(ldEmision, goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)

        Dim lcNumeroLote AS String = lcEmision & lcConsecutivo 
        Dim lcNumeroLoteSQL AS String = goServicios.mObtenerCampoFormatoSQL( lcNumeroLote )

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                                AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                                AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula                                 AS Cedula,")
            loConsulta.AppendLine("            CAST(" & lcConsecutivoSQL & " AS CHAR(2))           AS Consecutivo,")
            loConsulta.AppendLine("            CAST(" & lcEmisionSQL & " AS DATE)                  AS Emision,")
            loConsulta.AppendLine("            CAST(" & lcNumeroLoteSQL & " AS CHAR(8))            AS NumeroLote,")
            loConsulta.AppendLine("            (CASE WHEN COALESCE(Prop_Primer_Nombre.Val_Car, '') = ''")
            loConsulta.AppendLine("                THEN '[N/D]'")
            loConsulta.AppendLine("                ELSE Prop_Primer_Nombre.Val_Car")
            loConsulta.AppendLine("            END)                                                AS Primer_Nombre,")
            loConsulta.AppendLine("            (CASE WHEN COALESCE(Prop_Primer_Apellido.Val_Car, '') = ''")
            loConsulta.AppendLine("                THEN '[N/D]'")
            loConsulta.AppendLine("                ELSE Prop_Primer_Apellido.Val_Car")
            loConsulta.AppendLine("            END)                                                AS Primer_Apellido,")
            loConsulta.AppendLine("            (CASE WHEN COALESCE(Prop_Codigo_Interno.Val_Car, '') = ''")
            loConsulta.AppendLine("                THEN '[N/D]'")
            loConsulta.AppendLine("                ELSE Prop_Codigo_Interno.Val_Car")
            loConsulta.AppendLine("            END)                                                AS Codigo_Interno")
            loConsulta.AppendLine("FROM        Trabajadores")
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
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Codigo_Interno")
            loConsulta.AppendLine("        ON  Prop_Codigo_Interno.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Codigo_Interno.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Codigo_Interno.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Codigo_Interno.Cod_Pro = 'CODINTBONA'")
            loConsulta.AppendLine("WHERE   Trabajadores.Status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrden)
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoTXT_AlimentacionEmisionTEBCA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoTXT_AlimentacionEmisionTEBCA.ReportSource = loObjetoReporte

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

        '**************************************************
        ' Primero el registro de control: ENCABEZADO
        '**************************************************
        'Encabezado: #Fijo "0"
        loContenido.Append("0")

        'Encabezado: Número de Lote (8 caracteres)
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

        'Encabezado: #Fijo "1" (Emisión de tarjeta)
        loContenido.Append("1")

        'Encabezado: Fecha emisión (8 caracteres)
        Dim lcFecha As String = CDate(loRenglon("Emision")).ToString("yyyyMMdd")
        loContenido.Append(lcFecha)

        'Encabezado: (18 caracteres: relleno con "0")
        loContenido.Append("000000000000000000")

        'Fin de línea
        loContenido.Append(vbNewLine)

        '**************************************************
        ' Datos de trabajadores: MONTOS
        '**************************************************
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Datos: #Fijo "1" (Emisión de tarjeta)
            loContenido.Append("1")

            'Datos: Número de Lote (8 caracteres)
            loContenido.Append(lcLote)

            'Datos: Cédula
            Dim lcCedula As String = CStr(loRenglon("Cedula"))
            lcCedula = loLimpiaCedula.Replace(lcCedula, "")
            lcCedula = Strings.Left(lcCedula & Strings.Space(15), 15)
            loContenido.Append(lcCedula)

            'Datos: Primer Nombre
            Dim lcNombre As String = CStr(loRenglon("Primer_Nombre"))
            lcNombre = Strings.Left(lcNombre & Strings.Space(15), 15)
            loContenido.Append(lcNombre)

            'Datos: Primer Apellido
            Dim lcApellido As String = CStr(loRenglon("Primer_Apellido"))
            lcApellido = Strings.Left(lcApellido & Strings.Space(15), 15)
            loContenido.Append(lcApellido)

            'Datos: Código Interno
            Dim lcCodigoInt As String = CStr(loRenglon("Codigo_Interno"))
            lcCodigoInt = Strings.Left(lcCodigoInt & Strings.Space(15), 15)
            loContenido.Append(lcCodigoInt)
            
            If (n < lnCantidad-1) Then
                'Fin de línea: excepto en el último registro
                loContenido.Append(vbNewLine)
            End if
    
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
' RJG: 150/09/14: Código Inicial.
'-------------------------------------------------------------------------------------------'
