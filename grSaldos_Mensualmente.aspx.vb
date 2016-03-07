'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grSaldos_Mensualmente"
'-------------------------------------------------------------------------------------------'
Partial Class grSaldos_Mensualmente
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Construccion de la consulta
            '-------------------------------------------------------------------------------------------'

            ' Cálculo del saldo inicial por cada cuenta
            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("        (SUM(Movimientos_Cuentas.Mon_Deb)- SUM(Movimientos_Cuentas.Mon_Hab)) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO	#tempSALDOINICIAL     ")
            loComandoSeleccionar.AppendLine("FROM	Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("JOIN Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue     ")
            loComandoSeleccionar.AppendLine("JOIN Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban     ")
            loComandoSeleccionar.AppendLine("WHERE	 Movimientos_Cuentas.Fec_Ini < " & lcParametro2Desde & "  ")
            loComandoSeleccionar.AppendLine("     	 AND Movimientos_Cuentas.Cod_Con BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Movimientos_Cuentas.Cod_Cue     ")
            loComandoSeleccionar.AppendLine(" ")

            ' Listado de todos los movimientos por cuentas
            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("        Cuentas_Bancarias.Num_Cue,     ")
            loComandoSeleccionar.AppendLine("        Bancos.Nom_Ban,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Fec_Ini,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Documento,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Cod_Tip,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Tip_Doc,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Comentario,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Tip_Ori,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Deb,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Hab,     ")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Imp1     ")
            loComandoSeleccionar.AppendLine("INTO	#tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("JOIN Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue     ")
            loComandoSeleccionar.AppendLine("JOIN Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban     ")
            loComandoSeleccionar.AppendLine("WHERE	 Movimientos_Cuentas.Cod_Con BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Fec_Ini BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Cod_Tip BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado de movimiento por mes de cada cuenta
            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("        #tempMOVIMIENTO.Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("        #tempMOVIMIENTO.Num_Cue,     ")
            loComandoSeleccionar.AppendLine("        #tempMOVIMIENTO.Nom_Ban,     ")
            loComandoSeleccionar.AppendLine("        DATEPART(MONTH,#tempMOVIMIENTO.Fec_Ini) AS Mes,     ")
            loComandoSeleccionar.AppendLine("        DATEPART(YEAR,#tempMOVIMIENTO.Fec_Ini) AS Año,     ")
            loComandoSeleccionar.AppendLine("        SUM(#tempMOVIMIENTO.Mon_Deb) AS Mon_Deb,     ")
            loComandoSeleccionar.AppendLine("        SUM(#tempMOVIMIENTO.Mon_Hab) AS Mon_Hab,     ")
            loComandoSeleccionar.AppendLine("        SUM(#tempMOVIMIENTO.Mon_Imp1) AS Mon_Imp1,     ")
            loComandoSeleccionar.AppendLine("        0 AS Sal_Ini,     ")
            loComandoSeleccionar.AppendLine("        0 AS Sal_Doc     ")
            loComandoSeleccionar.AppendLine("INTO #tempRESULT")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("GROUP BY #tempMOVIMIENTO.Cod_Cue, #tempMOVIMIENTO.Num_Cue, #tempMOVIMIENTO.Nom_Ban, DATEPART(MONTH,#tempMOVIMIENTO.Fec_Ini), DATEPART(YEAR,#tempMOVIMIENTO.Fec_Ini)")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & ", DATEPART(YEAR,#tempMOVIMIENTO.Fec_Ini) ASC, DATEPART(MONTH,#tempMOVIMIENTO.Fec_Ini) ASC")
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

            ' Seleccion de las cuentas que han tenido movimiento
            loComandoSeleccionar.AppendLine("SELECT DISTINCT Movimientos_Cuentas.Cod_Cue INTO #tempCUENTASMOVIMIENTO FROM Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("")

            ' Tabla base para cada cuenta con las fechas y los montos en 0(Cero)
            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Cod_Cue,")
            loComandoSeleccionar.AppendLine(" 	        Cuentas_Bancarias.Num_Cue,")
            loComandoSeleccionar.AppendLine(" 	        Bancos.Nom_Ban,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.Mes,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.Año,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.mon_deb,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.mon_hab,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.mon_imp1,")
            loComandoSeleccionar.AppendLine(" 	        ISNULL(#tempSALDOINICIAL.sal_ini,0) AS sal_ini,")
            loComandoSeleccionar.AppendLine(" 	        #tempFECHAS.sal_doc")
            loComandoSeleccionar.AppendLine("INTO #tempBASICO")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Bancarias")
            loComandoSeleccionar.AppendLine("CROSS JOIN #tempFECHAS")
            loComandoSeleccionar.AppendLine("JOIN Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban     ")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tempSALDOINICIAL ON Cuentas_Bancarias.Cod_Cue = #tempSALDOINICIAL.Cod_Cue")
            loComandoSeleccionar.AppendLine("JOIN #tempCUENTASMOVIMIENTO ON Cuentas_Bancarias.Cod_Cue = #tempCUENTASMOVIMIENTO.Cod_Cue")
            loComandoSeleccionar.AppendLine("WHERE  	Cuentas_Bancarias.Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado con la union de la tabla base de cuentas con la tabla resultado de movimiento por cuentas
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		    #tempBASICO.Cod_Cue AS cod_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Num_Cue AS num_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Nom_Ban AS nom_ban,     ")
            loComandoSeleccionar.AppendLine("           (((DATEPART(YEAR," & lcParametro2Hasta & ") - #tempBASICO.año)*12)+(DATEPART(MONTH," & lcParametro2Hasta & ") - #tempBASICO.mes) + 1) AS num_meses,     ")
            loComandoSeleccionar.AppendLine("	        CASE")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempBASICO.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine("	        END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("           #tempBASICO.mes As mes,")
            loComandoSeleccionar.AppendLine("           #tempBASICO.año AS año,     ")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Mon_Deb AS mon_deb,     ")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Mon_Hab AS mon_hab,")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Mon_Imp1 AS mon_imp1,")
            loComandoSeleccionar.AppendLine("           #tempBASICO.Sal_Ini AS Sal_Ini,     ")
            loComandoSeleccionar.AppendLine("           #tempBASICO.sal_doc AS Sal_Doc  ")
            loComandoSeleccionar.AppendLine(" INTO #tempLISTRESULT")
            loComandoSeleccionar.AppendLine(" FROM #tempBASICO ")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Cod_Cue AS cod_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Num_Cue AS num_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Nom_Ban AS nom_ban,     ")
            loComandoSeleccionar.AppendLine("           (((DATEPART(YEAR," & lcParametro2Hasta & ") - #tempRESULT.año)*12)+(DATEPART(MONTH," & lcParametro2Hasta & ") - #tempRESULT.mes) + 1) AS num_meses,     ")
            loComandoSeleccionar.AppendLine("	        CASE")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("		        WHEN #tempRESULT.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine("	        END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("           #tempRESULT.mes As mes,")
            loComandoSeleccionar.AppendLine("           #tempRESULT.año AS año,     ")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Mon_Deb AS mon_deb,     ")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Mon_Hab AS mon_hab,")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Mon_Imp1 AS mon_imp1,")
            loComandoSeleccionar.AppendLine("           #tempRESULT.Sal_Ini AS Sal_Ini,     ")
            loComandoSeleccionar.AppendLine("           #tempRESULT.sal_doc AS Sal_Doc     ")
            loComandoSeleccionar.AppendLine(" FROM #tempRESULT")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla final de resultado
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.Cod_Cue AS cod_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.Num_Cue AS num_cue,     ")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.Nom_Ban AS nom_ban,     ")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.num_meses As num_meses,")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.str_mes As str_mes,")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.mes As mes,")
            loComandoSeleccionar.AppendLine("           #tempLISTRESULT.año AS año,")
            loComandoSeleccionar.AppendLine("           SUM(#tempLISTRESULT.Mon_Deb) AS mon_deb,     ")
            loComandoSeleccionar.AppendLine("           SUM(#tempLISTRESULT.Mon_Hab) AS mon_hab,")
            loComandoSeleccionar.AppendLine("           SUM(#tempLISTRESULT.Mon_Imp1) AS mon_imp1,")
            loComandoSeleccionar.AppendLine("           SUM(#tempLISTRESULT.Sal_Ini) AS Sal_Ini,     ")
            loComandoSeleccionar.AppendLine("           SUM(#tempLISTRESULT.sal_doc) AS Sal_Doc     ")
            loComandoSeleccionar.AppendLine(" FROM #tempLISTRESULT")
            loComandoSeleccionar.AppendLine(" GROUP BY #tempLISTRESULT.Cod_Cue,#tempLISTRESULT.Num_Cue,#tempLISTRESULT.Nom_Ban,#tempLISTRESULT.num_meses,#tempLISTRESULT.str_mes,#tempLISTRESULT.mes,#tempLISTRESULT.año")
            loComandoSeleccionar.AppendLine(" ORDER BY #tempLISTRESULT.Cod_Cue,#tempLISTRESULT.año,#tempLISTRESULT.mes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmetne los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Cue", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Num_Cue", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Ban", GetType(String))
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

                loNuevaFila.Item("Cod_Cue") = loFila("Cod_Cue")
                loNuevaFila.Item("Num_Cue") = loFila("Num_Cue")
                loNuevaFila.Item("Nom_Ban") = loFila("Nom_Ban")
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
                Cuenta_Actual = loFila("Cod_Cue")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Cue") <> Cuenta_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Cod_Cue") = loFila("Cod_Cue")
                    loNuevaFila.Item("Num_Cue") = loFila("Num_Cue")
                    loNuevaFila.Item("Nom_Ban") = loFila("Nom_Ban")
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
                    Cuenta_Actual = loFila("Cod_Cue")

                    loTabla.AcceptChanges()

                Next lnNumeroFila


                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'


                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grSaldos_Mensualmente", loDatosReporteFinal)
            Else
                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grSaldos_Mensualmente", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrSaldos_Mensualmente.ReportSource = loObjetoReporte

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
' DLC: 30/04/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' DLC: 19/07/2010: Ajuste de la consulta a la base de datos, acomodado el filtrado por 
'                   cuenta bancaria.
'-------------------------------------------------------------------------------------------'
' DLC: 06/09/2010: Ajuste de los nombres de las tablas temporales, asi como tambien,
'                   la documentacion de la consulta.
'-------------------------------------------------------------------------------------------'
