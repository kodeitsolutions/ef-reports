Imports System.Data
Partial Class rLibro_Ventas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Ini                                                                                                              AS  FechaCobro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des                                              AS  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec                                              AS  Por_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("           ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Net END),2)                            AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) END),2) AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Exe END),2)                            AS  Mon_Exe, ")
            loComandoSeleccionar.AppendLine("           ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Imp1 END),2)                           AS  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS  Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Por_Imp1 END),2)                           AS  Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN '03-Anu' ELSE '01-Reg' END)                                               AS  Status_Documento, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cuentas_Cobrar.Cod_Tip = 'RETIVA' THEN '01' ELSE '02' END)                                                       AS  Tipo_Documento ")
            loComandoSeleccionar.AppendLine("INTO		#tmpLibroVentas ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON  Cuentas_Cobrar.Cod_Cli	=   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cobros ON Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori")
            loComandoSeleccionar.AppendLine("						AND Renglones_Cobros.Cod_Tip   =   'RETIVA' ")
            loComandoSeleccionar.AppendLine("						AND Cuentas_Cobrar.Cod_Tip     =   Renglones_Cobros.Cod_Tip ")
            loComandoSeleccionar.AppendLine("	JOIN	Cobros ON Cobros.Documento           =   Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("						AND Cobros.Fec_Ini  BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("						AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Fec_Ini      < " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            
            
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            
            

            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("				(CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("				(CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des                                              AS  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec                                              AS  Por_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Net END),2)                            AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) END),2) AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Exe END),2)                            AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Imp1 END),2)                           AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE When Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Por_Imp1 END),2)                           AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN '03-Anu' ELSE '01-Reg' END)                                               AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'RETIVA' THEN '01' ELSE '02' END)                                                       AS Tipo_Documento ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'FACT' ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            
            

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            
            

            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("				(CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("				(CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Des * -1)                                   	AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Rec * -1)                                   	AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des                                              AS Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec                                              AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr1 * -1)                                   	AS Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr2 * -1)                                   	AS Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Otr3 * -1)                                   	AS Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Net * -1) END),2)                            AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE ((Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) * -1) END),2) AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Exe * -1) END),2)                            AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Imp1 * -1) END),2)                           AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Por_Imp1 END),2)                           AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN '03-Anu' ELSE '01-Reg' END)                                               AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'RETIVA' THEN '01' ELSE '02' END)                                                       AS Tipo_Documento ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            
            

            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des                                              AS  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec                                              AS  Por_Rec, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Net END),2)                            AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) END),2) AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Exe END),2)                            AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Mon_Imp1 END),2)                           AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			ROUND((CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN 0.00 ELSE Cuentas_Cobrar.Por_Imp1 End),2)                           AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado' THEN '03-Anu' ELSE '01-Reg' END)                                               AS Status_Documento, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Cod_Tip = 'RETIVA' THEN '01' ELSE '02' END)                                                       AS Tipo_Documento ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If 
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            

            loComandoSeleccionar.AppendLine("SELECT    * ")
            loComandoSeleccionar.AppendLine("FROM      #tmpLibroVentas ")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)

		    ' Me.mEscribirConsulta(lcComandoSeleccionar.ToString)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas.ReportSource = loObjetoReporte

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
' GMO: 02/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 15/11/08: Se acondiciono el select a la estructura actual de los reportes
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
' CMS:  22/05/10: Filtro Revision y Tipo de revision. Validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  05/09/11: Ajuste del Select y mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG:  21/11/11: Ajuste en la presentación de los descuentos/recargos: ahora se resta/suma	'
'				  proporcionalmente tanto ene l monto exento como en el grabable.			'
'-------------------------------------------------------------------------------------------'
