'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rTrabajadores_FAOV_CEGASA"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rTrabajadores_FAOV_CEGASA
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        'Dim lcParametro2Hasta As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
        'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        'Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        'Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        
        'Lista de conceptos considerados para calcular FAOV
        Dim laConceptosDevengadoFAOV As New ArrayList()
        Dim lcConceptosDevengadoFAOV As String

        laConceptosDevengadoFAOV.Add("'A000'")
        laConceptosDevengadoFAOV.Add("'A001'")
        laConceptosDevengadoFAOV.Add("'A002'")
        laConceptosDevengadoFAOV.Add("'A005'")
        laConceptosDevengadoFAOV.Add("'A030'")
        laConceptosDevengadoFAOV.Add("'A031'")
        laConceptosDevengadoFAOV.Add("'A032'")
        laConceptosDevengadoFAOV.Add("'A033'")
        laConceptosDevengadoFAOV.Add("'A034'")
        laConceptosDevengadoFAOV.Add("'A035'")
        laConceptosDevengadoFAOV.Add("'A036'")
        laConceptosDevengadoFAOV.Add("'A037'")
        laConceptosDevengadoFAOV.Add("'A038'")
        laConceptosDevengadoFAOV.Add("'A039'")
        laConceptosDevengadoFAOV.Add("'A040'")
        laConceptosDevengadoFAOV.Add("'A050'")
        laConceptosDevengadoFAOV.Add("'A051'")
        laConceptosDevengadoFAOV.Add("'A200'")
        laConceptosDevengadoFAOV.Add("'A300'")
        laConceptosDevengadoFAOV.Add("'A301'")
        laConceptosDevengadoFAOV.Add("'A302'")
        laConceptosDevengadoFAOV.Add("'A303'")
        laConceptosDevengadoFAOV.Add("'A304'")
        laConceptosDevengadoFAOV.Add("'A305'")
        laConceptosDevengadoFAOV.Add("'A320'")
        laConceptosDevengadoFAOV.Add("'A330'")
        laConceptosDevengadoFAOV.Add("'A402'")
        laConceptosDevengadoFAOV.Add("'A403'")
        laConceptosDevengadoFAOV.Add("'A404'")
        laConceptosDevengadoFAOV.Add("'A405'")
        laConceptosDevengadoFAOV.Add("'A406'")
        laConceptosDevengadoFAOV.Add("'A407'")
        laConceptosDevengadoFAOV.Add("'A501'")
        laConceptosDevengadoFAOV.Add("'A502'")
        laConceptosDevengadoFAOV.Add("'A504'")
        laConceptosDevengadoFAOV.Add("'A720'")
        laConceptosDevengadoFAOV.Add("'B001'")
        laConceptosDevengadoFAOV.Add("'B002'")
        laConceptosDevengadoFAOV.Add("'B004'")
        laConceptosDevengadoFAOV.Add("'B005'")
        laConceptosDevengadoFAOV.Add("'B007'")
        laConceptosDevengadoFAOV.Add("'B008'")
        laConceptosDevengadoFAOV.Add("'B010'")
        laConceptosDevengadoFAOV.Add("'B012'")
        laConceptosDevengadoFAOV.Add("'B013'")
        laConceptosDevengadoFAOV.Add("'B015'")
        laConceptosDevengadoFAOV.Add("'B016'")
        laConceptosDevengadoFAOV.Add("'B017'")
        laConceptosDevengadoFAOV.Add("'B020'")
        laConceptosDevengadoFAOV.Add("'B021'")
        laConceptosDevengadoFAOV.Add("'B030'")
        laConceptosDevengadoFAOV.Add("'B031'")
        laConceptosDevengadoFAOV.Add("'B032'")
        laConceptosDevengadoFAOV.Add("'B033'")
        laConceptosDevengadoFAOV.Add("'B060'")
        laConceptosDevengadoFAOV.Add("'B061'")
        laConceptosDevengadoFAOV.Add("'B100'")
        laConceptosDevengadoFAOV.Add("'B102'")
        laConceptosDevengadoFAOV.Add("'B500'")
        laConceptosDevengadoFAOV.Add("'B501'")
        laConceptosDevengadoFAOV.Add("'B600'")
        laConceptosDevengadoFAOV.Add("'B606'")

        lcConceptosDevengadoFAOV = String.Join(",", laConceptosDevengadoFAOV.ToArray())

        'Lista de conceptos que representan "deducción" en el aporte del Trabajador al FAOV en el Recibo
        Dim laConceptosDevengadoFAOV2 As New ArrayList()
        Dim lcConceptosDevengadoFAOV2 As String
        laConceptosDevengadoFAOV2.Add("'E001'")
        laConceptosDevengadoFAOV2.Add("'E005'")
        lcConceptosDevengadoFAOV2 = String.Join(",", laConceptosDevengadoFAOV2.ToArray())

        'Lista de conceptos que representan el aporte del Trabajador al FAOV en el Recibo
        Dim laConceptosFAOV_Trabajador As New ArrayList()
        Dim lcConceptosFAOV_Trabajador As String

        laConceptosFAOV_Trabajador.Add("'R003'")
        laConceptosFAOV_Trabajador.Add("'R303'")
        laConceptosFAOV_Trabajador.Add("'R403'")

        lcConceptosFAOV_Trabajador = String.Join(",", laConceptosFAOV_Trabajador.ToArray())

        'Lista de conceptos que representan el aporte del Patrono al FAOV en el Recibo
        Dim laConceptosFAOV_Patrono As New ArrayList()
        Dim lcConceptosFAOV_Patrono As String

        laConceptosFAOV_Patrono.Add("'U003'")
        laConceptosFAOV_Patrono.Add("'U303'")
        laConceptosFAOV_Patrono.Add("'U403'")

        lcConceptosFAOV_Patrono = String.Join(",", laConceptosFAOV_Patrono.ToArray())

        Try

            loConsulta.AppendLine("DECLARE @ldFecha AS DATETIME = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loConsulta.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(10) = " & lcParametro4Desde)
            loConsulta.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(10) = " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   ROW_NUMBER() OVER(ORDER BY Trabajadores.Cod_Tra ASC)    AS Numero,")
            loConsulta.AppendLine("         CAST(@ldFecha AS DATE)							        AS Periodo,")
            loConsulta.AppendLine("         COALESCE(Banavih.Val_Car,'00000000000000000000')        AS Numero_Cuenta_Banavih,")
            loConsulta.AppendLine("         Trabajadores.Cod_Tra                                    AS Cod_Tra,")
            loConsulta.AppendLine("         Trabajadores.Nom_Tra                                    AS Nom_Tra,")
            loConsulta.AppendLine("         Trabajadores.Fec_Ini                                    AS Fec_Ini,")
            loConsulta.AppendLine("         Trabajadores.Cedula                                     AS Cedula,")
            loConsulta.AppendLine("         COALESCE(Faov.Quincena1, 0)                             AS Quincena1,")
            loConsulta.AppendLine("         COALESCE(Faov.Quincena2, 0)                             AS Quincena2,")
            loConsulta.AppendLine("         COALESCE(Faov.Devengado, 0)                             AS Devengado,")
            loConsulta.AppendLine("         COALESCE(Faov.Aporte_Trabajador, 0)                     AS Aporte_Trabajador,")
            loConsulta.AppendLine("         COALESCE(Faov.Aporte_Patrono, 0)                        AS Aporte_Patrono,")
            loConsulta.AppendLine("         COALESCE(Faov.Aporte_Total, 0)                          AS Aporte_Total,")
            loConsulta.AppendLine("         (CASE WHEN COALESCE(Faov.Devengado, 0) > 0 ")
            loConsulta.AppendLine("             AND COALESCE(Faov.Aporte_Total, 0) = 0")
            loConsulta.AppendLine("             THEN 'NO APORTA'")
            loConsulta.AppendLine("             WHEN Trabajadores.Status = 'L' ")
            loConsulta.AppendLine("             THEN 'LIQUIDADO'")
            loConsulta.AppendLine("             ELSE COALESCE(Faov.Observaciones, '')")
            loConsulta.AppendLine("         END)                                                    AS Observaciones,")
            loConsulta.AppendLine("         COALESCE(Prop_Primer_Nombre.Val_Car, '')                AS Primer_Nombre,")
            loConsulta.AppendLine("         COALESCE(Prop_Segundo_Nombre.Val_Car, '')               AS Segundo_Nombre,")
            loConsulta.AppendLine("         COALESCE(Prop_Primer_Apellido.Val_Car, '')              AS Primer_Apellido,")
            loConsulta.AppendLine("         COALESCE(Prop_Segundo_Apellido.Val_Car, '')             AS Segundo_Apellido,")
            loConsulta.AppendLine("         (CASE WHEN Trabajadores.Status = 'L' ")
            loConsulta.AppendLine("             THEN COALESCE(Liquidacion.Fecha, Trabajadores.Fec_Fin)")
            loConsulta.AppendLine("             ELSE NULL ")
            loConsulta.AppendLine("         END)                                                    AS Fec_Fin")
            loConsulta.AppendLine("FROM Trabajadores")
            loConsulta.AppendLine("    LEFT JOIN(  SELECT  Recibos.Cod_Tra                          AS Cod_Tra,")
            loConsulta.AppendLine("                        MAX(CASE WHEN Recibos.Cod_Con='91' ")
            loConsulta.AppendLine("                                THEN 'Vacaciones'")
            loConsulta.AppendLine("                                ELSE ''")
            loConsulta.AppendLine("                            END)                                 AS Observaciones,")
            loConsulta.AppendLine("                        SUM(CASE WHEN DAY(Recibos.Fec_Fin)<16 AND Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV & ")")
            loConsulta.AppendLine("                                    THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                WHEN DAY(Recibos.Fec_Fin)<16 AND Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV2 & ")")
            loConsulta.AppendLine("                                    THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Quincena1,")
            loConsulta.AppendLine("                        SUM(CASE WHEN DAY(Recibos.Fec_Fin)>=16 AND Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV & ")")
            loConsulta.AppendLine("                                    THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                WHEN DAY(Recibos.Fec_Fin)>=16 AND Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV2 & ")")
            loConsulta.AppendLine("                                    THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Quincena2,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV & ")")
            loConsulta.AppendLine("                                    THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                WHEN Renglones_Recibos.Cod_Con IN (" & lcConceptosDevengadoFAOV2 & ")")
            loConsulta.AppendLine("                                    THEN -Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Devengado,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Cod_Con IN (" & lcConceptosFAOV_Trabajador & ")")
            loConsulta.AppendLine("                                THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Aporte_Trabajador,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Cod_Con IN (" & lcConceptosFAOV_Patrono & ")")
            loConsulta.AppendLine("                                THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Aporte_Patrono,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Cod_Con IN (" & lcConceptosFAOV_Trabajador & "," & lcConceptosFAOV_Patrono & ")")
            loConsulta.AppendLine("                                THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                                ELSE 0")
            loConsulta.AppendLine("                            END)                                 AS Aporte_Total                            ")
            loConsulta.AppendLine("                FROM Recibos")
            loConsulta.AppendLine("                    JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            loConsulta.AppendLine("                WHERE Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("                    AND MONTH(Recibos.fec_fin) = MONTH(CAST(@ldFecha AS DATE))")
            loConsulta.AppendLine("                    AND YEAR(Recibos.fec_fin) = YEAR(CAST(@ldFecha AS DATE))")
            loConsulta.AppendLine("                GROUP BY Recibos.Cod_Tra")
            loConsulta.AppendLine("            ) AS Faov")
            loConsulta.AppendLine("        ON  Faov.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Banavih")
            loConsulta.AppendLine("        ON  Banavih.Cod_Reg = " & goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo))
            loConsulta.AppendLine("        AND Banavih.Origen = 'Empresas'")
            loConsulta.AppendLine("        AND Banavih.Cod_Pro = 'NUMBANAVIH'")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Primer_Nombre")
            loConsulta.AppendLine("        ON  Prop_Primer_Nombre.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Primer_Nombre.Cod_Pro = 'NOMTRA01'")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Segundo_Nombre")
            loConsulta.AppendLine("        ON  Prop_Segundo_Nombre.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Segundo_Nombre.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Segundo_Nombre.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Segundo_Nombre.Cod_Pro = 'NOMTRA02'")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Primer_Apellido")
            loConsulta.AppendLine("        ON  Prop_Primer_Apellido.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Primer_Apellido.Cod_Pro = 'NOMTRA03'")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades Prop_Segundo_Apellido")
            loConsulta.AppendLine("        ON  Prop_Segundo_Apellido.Cod_Reg = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Prop_Segundo_Apellido.Origen = 'Trabajadores'")
            loConsulta.AppendLine("        AND Prop_Segundo_Apellido.Clase = 'Trabajador'")
            loConsulta.AppendLine("        AND Prop_Segundo_Apellido.Cod_Pro = 'NOMTRA04'")
            loConsulta.AppendLine("    LEFT JOIN (")
            loConsulta.AppendLine("                SELECT  Cod_Tra, MAX(Fecha) AS Fecha")
            loConsulta.AppendLine("                FROM    Liquidaciones")
            loConsulta.AppendLine("                WHERE   Liquidaciones.status = 'Confirmado'")
            loConsulta.AppendLine("                GROUP BY Cod_Tra")
            loConsulta.AppendLine("        ) Liquidacion")
            loConsulta.AppendLine("        ON Liquidacion.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("WHERE   Trabajadores.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine("        AND Trabajadores.Status IN ( " & lcParametro2Desde & " )")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            If lcParametro5Desde.ToUpper() = "NO" Then
                loConsulta.AppendLine("         AND COALESCE(Faov.Devengado, 0) > 0 ")
            End If
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rTrabajadores_FAOV_CEGASA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rTrabajadores_FAOV_CEGASA.ReportSource = loObjetoReporte

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
        
        Dim loRenglon As DataRow = loTabla.Rows(0)
        Dim ldFechaEmision As Date = CDate(loRenglon("Periodo"))
        Dim lcNumeroCuentaNominaBanavih As String = Cstr(loRenglon("Numero_Cuenta_Banavih"))
        Dim lcNombreArchivo As String = "N" & lcNumeroCuentaNominaBanavih & ldFechaEmision.ToString("MMyyyy")

        Dim loContenido As New StringBuilder()

        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        Dim lnCantidad As Integer = loTabla.Rows.Count
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Si el trabajador no aportó al periodo: no se incluye
            If (CDec(loRenglon("Aporte_Total")) <= 0) Then
                Continue For
            End If

            'Datos: Tipo de Identificación (debe ser V o E)
            Dim lcTipoId As String = CStr(loRenglon("Cedula")).ToUpper().Trim()
            If (lcTipoId.Length > 0) Then
                lcTipoId = Strings.Left(lcTipoId, 1)
            End If

            If (lcTipoId <> "E" AndAlso lcTipoId <> "V") Then
                lcTipoId = "" 'Con esto fallará al subir el TXT
            End If
            loContenido.Append(lcTipoId)
            loContenido.Append(",")

            'Datos: Cédula (hasta 10 caracteres numéricos, sin rellenar)
            Dim lcCedula As String = CStr(loRenglon("Cedula")).ToUpper()
            lcCedula = loLimpiaCedula.Replace(lcCedula, "")
            loContenido.Append(lcCedula)
            loContenido.Append(",")

            'Datos: Primer Nombre (hasta 25 caracteres, sin rellenar)
            Dim lcPrimerNombre As String = CStr(loRenglon("Primer_Nombre")).Trim()
            lcPrimerNombre = Strings.Left(lcPrimerNombre, 25)
            lcPrimerNombre = Me.mConvertirANSI(lcPrimerNombre)
            loContenido.Append(lcPrimerNombre)
            loContenido.Append(",")

            'Datos: Segundo Nombre (hasta 25 caracteres, sin rellenar)
            Dim lcSegundoNombre As String = CStr(loRenglon("Segundo_Nombre")).Trim()
            lcSegundoNombre = Strings.Left(lcSegundoNombre, 25)
            lcSegundoNombre = Me.mConvertirANSI(lcSegundoNombre)
            loContenido.Append(lcSegundoNombre)
            loContenido.Append(",")

            'Datos: Primer Apellido (hasta 25 caracteres, sin rellenar)
            Dim lcPrimerApellido As String = CStr(loRenglon("Primer_Apellido")).Trim()
            lcPrimerApellido = Strings.Left(lcPrimerApellido, 25)
            lcPrimerApellido = Me.mConvertirANSI(lcPrimerApellido)
            loContenido.Append(lcPrimerApellido)
            loContenido.Append(",")

            'Datos: Segundo Apellido (hasta 25 caracteres, sin rellenar)
            Dim lcSegundoApellido As String = CStr(loRenglon("Segundo_Apellido")).Trim()
            lcSegundoApellido = Strings.Left(lcSegundoApellido, 25)
            lcSegundoApellido = Me.mConvertirANSI(lcSegundoApellido)
            loContenido.Append(lcSegundoApellido)
            loContenido.Append(",")

            'Datos: Salario Integral (hasta 11 caracteres numéricos, los dos últimos son decimales.)
            Dim lnSalario As Long = CLng(CDec(loRenglon("Devengado"))*100)
            Dim lcSalario As String = lnSalario.ToString("00000000000")
            loContenido.Append(lcSalario)
            loContenido.Append(",")

            'Datos: Fecha de Ingreso (8 caracteres con formato DDMMYYYY)
            Dim ldIngreso As Date = CDate(loRenglon("Fec_Ini"))
            Dim lcIngreso As String = ldIngreso.ToString("ddMMyyyy")
            loContenido.Append(lcIngreso)
            loContenido.Append(",")

            'Datos: Fecha de Egreso (8 caracteres con formato DDMMYYYY, 
            'si el trabajador no está liquidado entonces se deja en blanco)
            If Not IsDBNull(loRenglon("Fec_Fin")) Then
                Dim ldEgreso As Date = CDate(loRenglon("Fec_Fin"))
                Dim lcEgreso As String = ldEgreso.ToString("ddMMyyyy")
                loContenido.Append(lcEgreso)
            End If
            
            If (n < lnCantidad-1) Then
                'Fin de línea: excepto en el último registro
                loContenido.Append(vbNewLine)
            End if
    
        Next n
        
        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcNombreArchivo & ".txt")
        Me.Response.ContentType = "text/plain"
        'Normalmente usamos UTF-8, pero BANAVIH espera codificación ANSI
        Me.Response.ContentEncoding = system.Text.Encoding.ASCII()
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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 09/06/15: Codigo inicial, a partir derTrabajadores_FAOV.                             '
'-------------------------------------------------------------------------------------------'
