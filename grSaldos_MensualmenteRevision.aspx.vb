'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grSaldos_MensualmenteRevision"
'-------------------------------------------------------------------------------------------'
Partial Class grSaldos_MensualmenteRevision
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Construccion de la consulta
            '-------------------------------------------------------------------------------------------'

            ' Cálculo del saldo inicial por cada cuenta
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       LTRIM(RTRIM(Movimientos_Cuentas.Cod_Rev)) AS Cod_Rev,")
            loComandoSeleccionar.AppendLine("       (SUM(Movimientos_Cuentas.Mon_Deb)- SUM(Movimientos_Cuentas.Mon_Hab)) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("INTO   #tempSALDOINICIAL")
            loComandoSeleccionar.AppendLine("FROM   Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("WHERE  Movimientos_Cuentas.Fec_Ini < " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Cod_Con BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND ((" & lcParametro4Desde & " = 'Igual' AND Movimientos_Cuentas.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & ")")
            loComandoSeleccionar.AppendLine("           OR (" & lcParametro4Desde & " <> 'Igual' AND Movimientos_Cuentas.Cod_Rev NOT BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "))")
            loComandoSeleccionar.AppendLine("GROUP BY LTRIM(RTRIM(Movimientos_Cuentas.Cod_Rev))")
            loComandoSeleccionar.AppendLine("")

            ' Se lista las revisiones
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       '' AS Cod_Rev,")
            loComandoSeleccionar.AppendLine("       'Sin Revisión' AS Nom_Rev")
            loComandoSeleccionar.AppendLine("INTO	#tempREVISIONESBASIC")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       Revisiones.Cod_Rev,")
            loComandoSeleccionar.AppendLine("       Revisiones.Nom_Rev")
            loComandoSeleccionar.AppendLine("FROM	Revisiones")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	DISTINCT")
            loComandoSeleccionar.AppendLine("		#tempREVISIONESBASIC.Cod_Rev,")
            loComandoSeleccionar.AppendLine("		#tempREVISIONESBASIC.Nom_Rev")
            loComandoSeleccionar.AppendLine("INTO	#tempREVISIONES")
            loComandoSeleccionar.AppendLine("FROM	#tempREVISIONESBASIC")
            loComandoSeleccionar.AppendLine("JOIN	Movimientos_Cuentas ON #tempREVISIONESBASIC.Cod_Rev = LTRIM(RTRIM(Movimientos_Cuentas.Cod_Rev))")
            loComandoSeleccionar.AppendLine("WHERE  Movimientos_Cuentas.Cod_Con BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND ((" & lcParametro4Desde & " = 'Igual' AND Movimientos_Cuentas.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & ")")
            loComandoSeleccionar.AppendLine("           OR (" & lcParametro4Desde & " <> 'Igual' AND Movimientos_Cuentas.Cod_Rev NOT BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "))")
            loComandoSeleccionar.AppendLine("")

            ' Tabla resultado de movimiento por mes de cada cuenta
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       #tempREVISIONES.Cod_Rev,")
            loComandoSeleccionar.AppendLine("       #tempREVISIONES.Nom_Rev,")
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH,Movimientos_Cuentas.Fec_Ini) AS Mes,")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR,Movimientos_Cuentas.Fec_Ini) AS Año,")
            loComandoSeleccionar.AppendLine("       SUM(Movimientos_Cuentas.Mon_Deb) AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("       SUM(Movimientos_Cuentas.Mon_Hab) AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("       SUM(Movimientos_Cuentas.Mon_Imp1) AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("       0 AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("       0 AS Sal_Doc")
            loComandoSeleccionar.AppendLine("INTO   #tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine("FROM   Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("JOIN	#tempREVISIONES ON LTRIM(RTRIM(Movimientos_Cuentas.Cod_Rev)) = #tempREVISIONES.Cod_Rev")
            loComandoSeleccionar.AppendLine("WHERE  Movimientos_Cuentas.Cod_Con BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Fec_Ini BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Movimientos_Cuentas.Cod_Tip BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND ((" & lcParametro4Desde & " = 'Igual' AND Movimientos_Cuentas.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & ")")
            loComandoSeleccionar.AppendLine("           OR (" & lcParametro4Desde & " <> 'Igual' AND Movimientos_Cuentas.Cod_Rev NOT BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "))")
            loComandoSeleccionar.AppendLine("GROUP BY #tempREVISIONES.Cod_Rev, #tempREVISIONES.Nom_Rev, DATEPART(MONTH,Movimientos_Cuentas.Fec_Ini), DATEPART(YEAR,Movimientos_Cuentas.Fec_Ini)")
            loComandoSeleccionar.AppendLine("ORDER BY #tempREVISIONES.Cod_Rev, DATEPART(YEAR,Movimientos_Cuentas.Fec_Ini) ASC, DATEPART(MONTH,Movimientos_Cuentas.Fec_Ini) ASC")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado base de fechas con montos 0(cero)
            Dim lcDateDesde As Date = New System.DateTime(Val(Mid(lcParametro2Desde, 2, 4)), Val(Mid(lcParametro2Desde, 6, 2)), Val(Mid(lcParametro2Desde, 8, 2)), 0, 0, 0)
            Dim lcDateHasta As Date = New System.DateTime(Val(Mid(lcParametro2Hasta, 2, 4)), Val(Mid(lcParametro2Hasta, 6, 2)), Val(Mid(lcParametro2Hasta, 8, 2)), 23, 59, 59)
            Dim lcNumMeses As Integer = ((Year(lcDateHasta) - Year(lcDateDesde)) * 12) + (Month(lcDateHasta) - Month(lcDateDesde)) + 1

            Dim lcMes As Integer = Month(lcDateDesde)
            Dim lcAño As Integer = Year(lcDateDesde)

            If lcNumMeses > 36 Then
                lcAño = Year(lcDateHasta) - 3
                lcMes = Month(lcDateHasta) + 1
                If lcMes = 13 Then
                    lcMes = 1
                    lcAño = lcAño + 1
                End If
                lcNumMeses = ((Year(lcDateHasta) - lcAño) * 12) + (Month(lcDateHasta) - lcMes) + 1
            End If

            loComandoSeleccionar.AppendLine("SELECT " & lcMes & " AS Mes, " & lcAño & " AS Año, 0 AS Mon_Deb, 0 AS Mon_Hab, 0 AS Mon_Imp1, 0 AS Sal_Ini, 0 AS Sal_Doc INTO #tempFECHAS")
            For lcIndex As Integer = 1 To lcNumMeses - 1
                lcMes = lcMes + 1
                If lcMes = 13 Then
                    lcMes = 1
                    lcAño = lcAño + 1
                End If
                loComandoSeleccionar.AppendLine("UNION ALL")
                loComandoSeleccionar.AppendLine("SELECT " & lcMes & " AS Mes, " & lcAño & " AS Año, 0 AS Mon_Deb, 0 AS Mon_Hab, 0 AS Mon_Imp1, 0 AS Sal_Ini, 0 AS Sal_Doc")
            Next lcIndex
            loComandoSeleccionar.AppendLine("")

            ' Tabla resultado con la union de la tabla base de cuentas con la tabla resultado de movimiento por cuentas
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       #tempREVISIONES.Cod_Rev AS cod_rev,")
            loComandoSeleccionar.AppendLine("       #tempREVISIONES.Nom_Rev AS nom_rev,")
            loComandoSeleccionar.AppendLine("       (((DATEPART(YEAR," & lcParametro2Hasta & ") - #tempFECHAS.año)*12)+(DATEPART(MONTH," & lcParametro2Hasta & ") - #tempFECHAS.mes) + 1) AS num_meses,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tempFECHAS.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine("       END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.mes As mes,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.año AS año,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.Mon_Deb AS mon_deb,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.Mon_Hab AS mon_hab,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.Mon_Imp1 AS mon_imp1,")
            loComandoSeleccionar.AppendLine("       ISNULL(#tempSALDOINICIAL.sal_ini,0) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("       #tempFECHAS.sal_doc AS Sal_Doc")
            loComandoSeleccionar.AppendLine("INTO #tempLISTRESULT")
            loComandoSeleccionar.AppendLine("FROM #tempREVISIONES")
            loComandoSeleccionar.AppendLine("CROSS JOIN #tempFECHAS")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tempSALDOINICIAL ON #tempREVISIONES.Cod_Rev = #tempSALDOINICIAL.Cod_Rev")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Cod_Rev AS cod_rev,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Nom_Rev AS nom_rev,")
            loComandoSeleccionar.AppendLine("       (((DATEPART(YEAR," & lcParametro2Hasta & ") - #tempMOVIMIENTO.año)*12)+(DATEPART(MONTH," & lcParametro2Hasta & ") - #tempMOVIMIENTO.mes) + 1) AS num_meses,     ")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tempMOVIMIENTO.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine("       END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.mes As mes,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.año AS año,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Mon_Deb AS mon_deb,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Mon_Hab AS mon_hab,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Mon_Imp1 AS mon_imp1,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.Sal_Ini AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("       #tempMOVIMIENTO.sal_doc AS Sal_Doc")
            loComandoSeleccionar.AppendLine("FROM #tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla final de resultado
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tempLISTRESULT.Cod_Rev='' THEN 'SIN_REV' ELSE #tempLISTRESULT.Cod_Rev END AS cod_rev,")
            loComandoSeleccionar.AppendLine("       #tempLISTRESULT.Nom_Rev AS nom_rev,")
            loComandoSeleccionar.AppendLine("       #tempLISTRESULT.num_meses As num_meses,")
            loComandoSeleccionar.AppendLine("       #tempLISTRESULT.str_mes As str_mes,")
            loComandoSeleccionar.AppendLine("       #tempLISTRESULT.mes As mes,")
            loComandoSeleccionar.AppendLine("       #tempLISTRESULT.año AS año,")
            loComandoSeleccionar.AppendLine("       SUM(#tempLISTRESULT.Mon_Deb) AS mon_deb,")
            loComandoSeleccionar.AppendLine("       SUM(#tempLISTRESULT.Mon_Hab) AS mon_hab,")
            loComandoSeleccionar.AppendLine("       SUM(#tempLISTRESULT.Mon_Imp1) AS mon_imp1,")
            loComandoSeleccionar.AppendLine("       SUM(#tempLISTRESULT.Sal_Ini) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("       SUM(#tempLISTRESULT.sal_doc) AS Sal_Doc")
            loComandoSeleccionar.AppendLine("FROM #tempLISTRESULT")
            loComandoSeleccionar.AppendLine("GROUP BY #tempLISTRESULT.Cod_Rev,#tempLISTRESULT.Nom_Rev,#tempLISTRESULT.num_meses,#tempLISTRESULT.str_mes,#tempLISTRESULT.mes,#tempLISTRESULT.año")
            loComandoSeleccionar.AppendLine("ORDER BY #tempLISTRESULT.Cod_Rev,#tempLISTRESULT.año,#tempLISTRESULT.mes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmetne los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Rev", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Rev", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Num_Meses", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Str_Mes", GetType(String))
                loColumna.MaxLength = 4
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mes", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Año", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Imp1", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Cuenta_Actual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loNuevaFila.Item("Cod_Rev") = loFila("Cod_Rev")
                loNuevaFila.Item("Nom_Rev") = loFila("Nom_Rev")
                loNuevaFila.Item("Num_Meses") = loFila("Num_Meses")
                loNuevaFila.Item("Str_Mes") = loFila("Str_Mes")
                loNuevaFila.Item("Mes") = loFila("Mes")
                loNuevaFila.Item("Año") = loFila("Año")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Mon_Imp1") = loFila("Mon_Imp1")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Cuenta_Actual = loFila("Cod_Rev")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Rev") <> Cuenta_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Cod_Rev") = loFila("Cod_Rev")
                    loNuevaFila.Item("Nom_Rev") = loFila("Nom_Rev")
                    loNuevaFila.Item("Num_Meses") = loFila("Num_Meses")
                    loNuevaFila.Item("Str_Mes") = loFila("Str_Mes")
                    loNuevaFila.Item("Mes") = loFila("Mes")
                    loNuevaFila.Item("Año") = loFila("Año")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Mon_Imp1") = loFila("Mon_Imp1")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Cuenta_Actual = loFila("Cod_Rev")

                    loTabla.AcceptChanges()

                Next lnNumeroFila


                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'


                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grSaldos_MensualmenteRevision", loDatosReporteFinal)
            Else
                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grSaldos_MensualmenteRevision", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrSaldos_MensualmenteRevision.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' DLC: 06/09/2010: Programacion inicial (Tomando como base el reporte "Grafico de Saldos Mensualmente")
'-------------------------------------------------------------------------------------------'
