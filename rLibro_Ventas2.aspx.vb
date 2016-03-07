Imports System.Data

Partial Class rLibro_Ventas2 
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
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT    'Cobros' AS Tabla, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cobros.Fec_Ini                                                                                                              AS  FechaCobro, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento AS Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Net End),2)                            AS  Mon_Net, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) End),2) AS  Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Exe End),2)                            AS  Mon_Exe, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Imp1 End),2)                           AS  Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS  Cod_Imp, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Por_Imp1 End),2)                           AS  Por_Imp1, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS  Status_Documento, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Cod_Tip = 'RETIVA' Then '01' Else '02' End)                                                       AS  Tipo_Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Status												AS  Status, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal1                                              AS  Fiscal1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal2                                              AS  Fiscal2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal3                                              AS  Fiscal3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal4                                              AS  Fiscal4, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal5                                              AS  Fiscal5, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Referencia                                             AS  Referencia ")
            lcComandoSeleccionar.AppendLine(" INTO      #tmpLibroVentas ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            lcComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros ON Cuentas_Cobrar.Documento		=   Renglones_Cobros.Doc_Ori")
            lcComandoSeleccionar.AppendLine("							AND Renglones_Cobros.Cod_Tip    =   'RETIVA' ")
            lcComandoSeleccionar.AppendLine("							AND Cuentas_Cobrar.Cod_Tip     =   Renglones_Cobros.Cod_Tip ")
            lcComandoSeleccionar.AppendLine(" JOIN Cobros ON Cobros.Documento            =   Renglones_Cobros.Documento")
            lcComandoSeleccionar.AppendLine("           And Cobros.Fec_Ini              BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Fec_Ini      < " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status        IN ( " & lcParametro6Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev between " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("")
            

            lcComandoSeleccionar.AppendLine(" UNION ALL ")
            
            
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine(" SELECT    'Cobros' AS Tabla, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento AS Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Net End),2)                            AS Mon_Net, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) End),2) AS Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Exe End),2)                            AS Mon_Exe, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Imp1 End),2)                           AS Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Por_Imp1 End),2)                           AS Por_Imp1, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS Status_Documento, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Cod_Tip = 'RETIVA' Then '01' Else '02' End)                                                       AS Tipo_Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Status												AS  Status, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal1                                              AS  Fiscal1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal2                                              AS  Fiscal2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal3                                              AS  Fiscal3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal4                                              AS  Fiscal4, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal5                                              AS  Fiscal5, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Referencia                                           AS  Referencia ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Tip     =   'FACT' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status        IN ( " & lcParametro6Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev between " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("")
            

            lcComandoSeleccionar.AppendLine(" UNION ALL ")
            
            
            
        	lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine(" SELECT    'Cobros' AS Tabla, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento AS Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Net * -1) End),2)                            AS Mon_Net, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else ((Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) * -1) End),2) AS Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Exe * -1) End),2)                            AS Mon_Exe, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Imp1 * -1) End),2)                           AS Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Por_Imp1 End),2)                           AS Por_Imp1, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS Status_Documento, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Cod_Tip = 'RETIVA' Then '01' Else '02' End)                                                       AS Tipo_Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Status												AS  Status, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal1                                              AS  Fiscal1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal2                                              AS  Fiscal2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal3                                              AS  Fiscal3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal4                                              AS  Fiscal4, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal5                                              AS  Fiscal5, ")
            lcComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Des * -1)                                       AS  Mon_Des, ")
            lcComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Rec * -1)                                       AS  Mon_Rec, ")
            lcComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Otr1 * -1)                                      AS  Mon_Otr1, ")
            lcComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Otr2 * -1)                                      AS  Mon_Otr2, ")
            lcComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Otr3 * -1)                                      AS  Mon_Otr3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Referencia                                           AS  Referencia ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status        IN ( " & lcParametro6Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev between " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("")
            

            lcComandoSeleccionar.AppendLine(" UNION ALL ")


			lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine(" SELECT    'Cobros' AS Tabla, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento AS Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Net End),2)                            AS Mon_Net, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) End),2) AS Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Exe End),2)                            AS Mon_Exe, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Imp1 End),2)                           AS Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS Cod_Imp, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Por_Imp1 End),2)                           AS Por_Imp1, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS Status_Documento, ")
            lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Cod_Tip = 'RETIVA' Then '01' Else '02' End)                                                       AS Tipo_Documento, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Status												AS  Status, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal1                                              AS  Fiscal1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal2                                              AS  Fiscal2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal3                                              AS  Fiscal3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal4                                              AS  Fiscal4, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fiscal5                                              AS  Fiscal5, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des                                              AS  Mon_Des, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec                                              AS  Mon_Rec, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr1                                             AS  Mon_Otr1, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr2                                             AS  Mon_Otr2, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Otr3                                             AS  Mon_Otr3, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Referencia                                           AS  Referencia ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes ON (Cuentas_Cobrar.Cod_Cli  =   Clientes.Cod_Cli )")
            lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status        IN ( " & lcParametro6Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev between " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("")
            
            Dim lcStatus AS String = "Pendiente,Confirmado,Procesado,Pagado,Cerrado,Afectado,Serializado,Contabilizado,Iniciado,Conciliado,Otro,Anulado"
			
			If (cusAplicacion.goReportes.paParametrosIniciales(6).Equals(lcStatus)) Then

					lcComandoSeleccionar.AppendLine("UNION ALL")
			End If
			
			If (cusAplicacion.goReportes.paParametrosIniciales(6).Equals("Anulado")) Then

					lcComandoSeleccionar.AppendLine("UNION ALL")
			End If
			
			lcComandoSeleccionar.AppendLine("")
			lcComandoSeleccionar.AppendLine("")
			lcComandoSeleccionar.AppendLine("")
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''Obtencion de las Facturas de Venta Anuladas '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			 lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine(" SELECT    'Ventas' AS Tabla,")
            lcComandoSeleccionar.AppendLine("           Facturas.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fec_Ini                                                                                                      AS  FechaCobro, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           'FACT' AS Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Documento AS Documento, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Control, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Doc_Ori, ")
            lcComandoSeleccionar.AppendLine("           ROUND(Facturas.Mon_Net,2)								AS Mon_Net, ")
            lcComandoSeleccionar.AppendLine("           ROUND((Facturas.Mon_Bru - Facturas.Mon_Exe),2)			AS Mon_Bru, ")
            lcComandoSeleccionar.AppendLine("           ROUND(Facturas.Mon_Exe,2)								AS Mon_Exe, ")
            lcComandoSeleccionar.AppendLine("           ROUND(Facturas.Mon_Imp1,2)							AS Mon_Imp1, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Cod_Imp1                                                                                                      AS Cod_Imp, ")
            lcComandoSeleccionar.AppendLine("           ROUND(Facturas.Por_Imp1,2)                           AS Por_Imp1, ")
            lcComandoSeleccionar.AppendLine("           (Case When Facturas.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS Status_Documento, ")
            lcComandoSeleccionar.AppendLine("           '03'															AS Tipo_Documento, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Status													AS  Status, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fiscal1												AS  Fiscal1, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fiscal2												AS  Fiscal2, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fiscal3												AS  Fiscal3, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fiscal4												AS  Fiscal4, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Fiscal5												AS  Fiscal5, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Mon_Des1												AS  Mon_Des, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Mon_Rec1												AS  Mon_Rec, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Mon_Otr1												AS  Mon_Otr1, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Mon_Otr2												AS  Mon_Otr2, ")
            lcComandoSeleccionar.AppendLine("           Facturas.Mon_Otr3												AS  Mon_Otr3, ")
            lcComandoSeleccionar.AppendLine("           ''																AS  Referencia ")
            lcComandoSeleccionar.AppendLine(" FROM      Facturas ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes ON (Facturas.Cod_Cli  =   Clientes.Cod_Cli )")
            lcComandoSeleccionar.AppendLine(" WHERE     Facturas.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Facturas.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Facturas.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Facturas.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("			AND Facturas.Status  = 'Anulado'")
            If lcParametro5Desde = "Igual" Then
                lcComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Rev between " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Rev NOT between " & lcParametro4Desde)
            End If
            lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("")

            lcComandoSeleccionar.AppendLine(" SELECT    * ")
            lcComandoSeleccionar.AppendLine(" FROM      #tmpLibroVentas ")
            lcComandoSeleccionar.AppendLine(" ORDER BY Tabla ASC," & lcOrdenamiento)

            
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")
            
            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())
            
            Dim loTabla As New DataTable()
            
            If (cusAplicacion.goReportes.paParametrosIniciales(6).Equals(lcStatus))	 OR	 (cusAplicacion.goReportes.paParametrosIniciales(6).Equals("Anulado")) Then
					
							loTabla = laDatosReporte.Tables(0)
							
			Else
							loTabla = laDatosReporte.Tables(1)
							
			End If
			
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
			
			Dim laDatosReporte2 AS New DataSet()
			Dim loTabla2 AS New DataTable()
			
			loTabla2= loTabla.Copy()
			
			laDatosReporte2.Tables.Add(loTabla2)
			
			laDatosReporte2.AcceptChanges()
			
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas2", laDatosReporte2)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas2.ReportSource = loObjetoReporte

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
' MAT:  02/05/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 12/05/11: Ajuste para que muestre Sección de Facturas Anuladas
'-------------------------------------------------------------------------------------------'
