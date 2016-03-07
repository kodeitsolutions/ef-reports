'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMovimientosBancos_TMensual"
'-------------------------------------------------------------------------------------------'
Partial Class rMovimientosBancos_TMensual
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosFinales(6)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("  SELECT   ")
            loComandoSeleccionar.AppendLine("              Cuentas_Bancarias.Cod_Cue,   ")
            loComandoSeleccionar.AppendLine("              Cuentas_Bancarias.Nom_Cue,   ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 1 THEN  Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Ene_Mon_Deb,     ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 2 THEN Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Feb_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 3 THEN Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Mar_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 4 THEN  Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Abr_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 5 THEN  Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as May_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 6 THEN  Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Jun_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 7 THEN  Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Jul_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 8 THEN Movimientos_Cuentas.Mon_Deb     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Ago_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 9 THEN Movimientos_Cuentas.Mon_Deb   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Sep_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 10 THEN  Movimientos_Cuentas.Mon_Deb   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Oct_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 11 THEN  Movimientos_Cuentas.Mon_Deb   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Nov_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 12 THEN  Movimientos_Cuentas.Mon_Deb   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Dic_Mon_Deb,   ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 1 THEN  Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Ene_Mon_Hab,     ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 2 THEN Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Feb_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 3 THEN Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Mar_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 4 THEN  Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Abr_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 5 THEN  Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as May_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 6 THEN  Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Jun_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 7 THEN  Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Jul_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 8 THEN Movimientos_Cuentas.Mon_Hab     ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Ago_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 9 THEN Movimientos_Cuentas.Mon_Hab   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Sep_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 10 THEN  Movimientos_Cuentas.Mon_Hab   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Oct_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 11 THEN  Movimientos_Cuentas.Mon_Hab   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Nov_Mon_Hab,      ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 12 THEN  Movimientos_Cuentas.Mon_Hab   ")
            loComandoSeleccionar.AppendLine("  	            else 0      ")
            loComandoSeleccionar.AppendLine("              END as Dic_Mon_Hab,   ")
            loComandoSeleccionar.AppendLine("              CASE       ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 1 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)    ")
            loComandoSeleccionar.AppendLine("  	            else 0        ")
            loComandoSeleccionar.AppendLine("              END as Ene_Dif,      ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 2 THEN (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Feb_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 3 THEN (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Mar_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 4 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Abr_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 5 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as May_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 6 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Jun_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 7 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Jul_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 8 THEN (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)      ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Ago_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 9 THEN (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)    ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Sep_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 10 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)    ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Oct_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 11 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)    ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Nov_Dif,       ")
            loComandoSeleccionar.AppendLine("              CASE        ")
            loComandoSeleccionar.AppendLine("  	            WHEN DATEPART(MONTH, Movimientos_Cuentas.Fec_ini)  = 12 THEN  (Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)    ")
            loComandoSeleccionar.AppendLine("  	            else 0       ")
            loComandoSeleccionar.AppendLine("              END as Dic_Dif,    ")
            loComandoSeleccionar.AppendLine("              Movimientos_Cuentas.Mon_Deb AS Total_Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("              Movimientos_Cuentas.Mon_Hab AS Total_Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("              Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab AS Total_Dif    ")
            loComandoSeleccionar.AppendLine("  INTO	#tmpTemporal    ")
            loComandoSeleccionar.AppendLine("  FROM    ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas,    ")
            loComandoSeleccionar.AppendLine("          Cuentas_Bancarias    ")
            loComandoSeleccionar.AppendLine("  WHERE ")
            loComandoSeleccionar.AppendLine("  Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue    ")

            loComandoSeleccionar.AppendLine("             			AND Movimientos_Cuentas.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             			AND Movimientos_Cuentas.Cod_Cue between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             			AND Movimientos_Cuentas.Cod_Mon between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         				AND Movimientos_Cuentas.status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("             			AND Movimientos_Cuentas.Cod_Con between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         				AND " & lcParametro4Hasta)

            If lcParametro6Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Rev between " & lcParametro5Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Movimientos_Cuentas.Cod_Rev NOT between " & lcParametro5Desde)
            End If

            loComandoSeleccionar.AppendLine("         				AND " & lcParametro5Hasta)


            loComandoSeleccionar.AppendLine("  SELECT        ")
            loComandoSeleccionar.AppendLine("          #tmpTemporal.Cod_Cue,			    ")
            loComandoSeleccionar.AppendLine("          #tmpTemporal.Nom_Cue,			    ")
            loComandoSeleccionar.AppendLine("          sum(Ene_Mon_Deb) as Ene_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Feb_Mon_Deb) as Feb_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Mar_Mon_Deb) as Mar_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Abr_Mon_Deb) as Abr_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(May_Mon_Deb) as May_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Jun_Mon_Deb) as Jun_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Jul_Mon_Deb) as Jul_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Ago_Mon_Deb) as Ago_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Sep_Mon_Deb) as Sep_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Oct_Mon_Deb) as Oct_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Nov_Mon_Deb) as Nov_Mon_Deb,       ")
            loComandoSeleccionar.AppendLine("          sum(Dic_Mon_Deb) as Dic_Mon_Deb,      ")
            loComandoSeleccionar.AppendLine("          sum(Ene_Mon_Hab) as Ene_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Feb_Mon_Hab) as Feb_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Mar_Mon_Hab) as Mar_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Abr_Mon_Hab) as Abr_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(May_Mon_Hab) as May_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Jun_Mon_Hab) as Jun_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Jul_Mon_Hab) as Jul_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Ago_Mon_Hab) as Ago_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Sep_Mon_Hab) as Sep_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Oct_Mon_Hab) as Oct_Mon_Hab,       ")
		    loComandoSeleccionar.AppendLine("          sum(Nov_Mon_Hab) as Nov_Mon_Hab,       ")
            loComandoSeleccionar.AppendLine("          sum(Dic_Mon_Hab) as Dic_Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("          sum(Ene_Dif) as Ene_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Feb_Dif) as Feb_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Mar_Dif) as Mar_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Abr_Dif) as Abr_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(May_Dif) as May_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Jun_Dif) as Jun_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Jul_Dif) as Jul_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Ago_Dif) as Ago_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Sep_Dif) as Sep_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Oct_Dif) as Oct_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Nov_Dif) as Nov_Dif,       ")
            loComandoSeleccionar.AppendLine("          sum(Dic_Dif) as Dic_Dif,    ")
            loComandoSeleccionar.AppendLine("          Sum(Total_Mon_Deb) as Total_Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("          sum(Total_Mon_Hab) as Total_Mon_HAb,    ")
            loComandoSeleccionar.AppendLine("          sum(Total_Dif) as Total_Dif    ")
            loComandoSeleccionar.AppendLine("  FROM	#tmpTemporal       ")
            loComandoSeleccionar.AppendLine("  Group By    #tmpTemporal.Cod_Cue, #tmpTemporal.Nom_Cue    ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMovimientosBancos_TMensual", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMovimientosBancos_TMensual.ReportSource = loObjetoReporte


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
' CMS:  22/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 31/07/09: Filtro "Revision:", Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATENAME por DATEPART
'-------------------------------------------------------------------------------------------'