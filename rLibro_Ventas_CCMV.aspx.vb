'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'Imports Microsoft.Office.Interop

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_CCMV

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Dim loAppExcel As Excel.Application

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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cobros.Fec_Ini                                                                                                              AS  FechaCobro, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")

            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
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

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Rev, ")
            'lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" INTO      #tmpLibroVentas ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            lcComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros on (Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip And Cuentas_Cobrar.Documento    = Renglones_Cobros.Doc_Ori) ")
            lcComandoSeleccionar.AppendLine(" JOIN Cobros on Cobros.Documento = Renglones_Cobros.Documento")
            'lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine(" WHERE      ")

            lcComandoSeleccionar.AppendLine("           Cobros.Documento            =   Renglones_Cobros.Documento ")
            lcComandoSeleccionar.AppendLine("           And (Cuentas_Cobrar.Cod_Tip     =   Renglones_Cobros.Cod_Tip ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori) ")
            lcComandoSeleccionar.AppendLine("           And Renglones_Cobros.Cod_Tip    =   'RETIVA' ")
            lcComandoSeleccionar.AppendLine("           And Cobros.Fec_Ini              BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      < " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)

            'lcComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            'lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            'lcComandoSeleccionar.AppendLine("           Cobros.Fec_Ini                                                                                                              AS  FechaCobro, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'lcComandoSeleccionar.AppendLine("           Renglones_Cobros.Cod_Tip, ")
            'lcComandoSeleccionar.AppendLine("           Renglones_Cobros.Doc_Ori                                                                                                    AS  Documento, ")
            'lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            'lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            'lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Net End),2)                            AS  Mon_Net, ")
            'lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) End),2) AS  Mon_Bru, ")
            'lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Exe End),2)                            AS  Mon_Exe, ")
            'lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Mon_Imp1 End),2)                           AS  Mon_Imp1, ")
            'lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Imp                                                                                                      AS  Cod_Imp, ")
            'lcComandoSeleccionar.AppendLine("           ROUND((Case When Cuentas_Cobrar.Status = 'Anulado' Then 0.00 Else Cuentas_Cobrar.Por_Imp1 End),2)                           AS  Por_Imp1, ")
            'lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Status = 'Anulado' Then '03-Anu' Else '01-Reg' End)                                               AS  Status_Documento, ")
            'lcComandoSeleccionar.AppendLine("           (Case When Cuentas_Cobrar.Cod_Tip = 'RETIVA' Then '01' Else '02' End)                                                       AS  Tipo_Documento ")
            'lcComandoSeleccionar.AppendLine(" INTO      #tmpRetencionesDentroPeriodo ")
            'lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar, ")
            'lcComandoSeleccionar.AppendLine("           Clientes, ")
            'lcComandoSeleccionar.AppendLine("           Cobros, ")
            'lcComandoSeleccionar.AppendLine("           Renglones_Cobros ")
            'lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            'lcComandoSeleccionar.AppendLine("           And Cobros.Documento            =   Renglones_Cobros.Documento ")
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      =   Renglones_Cobros.Cod_Tip ")
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori ")
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      =   'RETIVA' ")
            'lcComandoSeleccionar.AppendLine("           And Cobros.Fec_Ini              BETWEEN " & lcParametro0Desde)
            'lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      >=  " & lcParametro0Desde)
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            'lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            'lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            'lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)

            lcComandoSeleccionar.AppendLine(" UNION ALL ")

            lcComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")

            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
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

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Rev, ")
            'lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN  Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            'lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine(" WHERE     ")

            
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip     =   'FACT' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)

            lcComandoSeleccionar.AppendLine(" UNION ALL ")

            lcComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")

            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
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

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Rev, ")
            'lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli")
            'lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine(" WHERE     ")
            

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)

            lcComandoSeleccionar.AppendLine(" UNION ALL ")

            lcComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(ORDER BY Cuentas_Cobrar.Fec_Ini) As Renglon, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  Fec_Ini, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini                                                                                                      AS  FechaCobro, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'lcComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")

            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            lcComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            lcComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
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

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Rev, ")
            'lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli")
            'lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)

            lcComandoSeleccionar.AppendLine(" SELECT    #tmpLibroVentas.*, ")
            lcComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Fec_Ini, 112)	AS	Fecha2	")
            lcComandoSeleccionar.AppendLine(" INTO      #tmpLibroVentas002 ")
            lcComandoSeleccionar.AppendLine(" FROM      #tmpLibroVentas ")

            lcComandoSeleccionar.AppendLine(" SELECT    * ")
            lcComandoSeleccionar.AppendLine(" FROM      #tmpLibroVentas002 ")
            'lcComandoSeleccionar.AppendLine(" ORDER BY  Fecha2, Documento ")
            lcComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(lcComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString)



            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLibro_Ventas_CCMV.ReportSource = loObjetoReporte

            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rLibro_Ventas_CCMV_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.modificar_excel(lcFileName, laDatosReporte)

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rLibro_Ventas_CCMV.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()
                
				Me.Response.End()
                
            End If

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                      "No se pudo Completar el Proceso: " & loExcepcion.Message & ", StackTrace: " & loExcepcion.StackTrace, _
                       vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                       "600px", _
                       "500px")

        End Try

    End Sub
    
''' <summary>
''' Rutina que modifica y formatea el contenido del archivo excel a descargar.
''' </summary>
''' <param name="loFileName">Ruta del archivo a modificar.</param>
''' <param name="loDatosReporte">Datos de la consulta a la base de datos.</param>
''' <remarks></remarks>
    Private Sub modificar_excel(ByVal loFileName As String, ByVal loDatosReporte As DataSet)

        Try
            ' Se inicializa el objeto a la aplicacion excel
            loAppExcel = New Excel.Application()
            loAppExcel.Visible = False
            loAppExcel.DisplayAlerts = False 

            ' Se carga el archivo excel a modificar
            Dim lcLibrosExcel As Excel.Workbooks = loAppExcel.Workbooks
            Dim lcLibroExcel As Excel.Workbook = lcLibrosExcel.Open(loFileName)

            ' Se activa la primera hoja del libro donde se almacenara toda la informacion
            Dim lcHojaExcel As Excel.Worksheet
            lcHojaExcel = lcLibroExcel.Worksheets(1)
            lcHojaExcel.Activate()

            ' Se selecciona toda la hoja para blanquera todo
            '   - El número total de columnas disponibles en Excel
            '       Viejo límite: 256 (2^8)         (Excel 2003 o inferor)
            '       Nuevo límite: 16.384 (2^14)     (Excel 2007)
            '   - El número total de filas disponibles en Excel
            '       Viejo límite: 65.536 (2^16)         (Excel 2003 o inferior)
            '       Nuevo límite: el 1.048.576 (2^20)   (Excel 2007)
            Dim lcRango As Excel.Range = lcHojaExcel.Range("A1:IV65536")
            lcRango.Select()
            lcRango.Clear()
            lcRango.Font.Size = 9
            lcRango.Font.Name = "Tahoma"



            ' Nombre de la empresa
            lcHojaExcel.Cells(1, 1).Value = cusAplicacion.goEmpresa.pcNombre.ToUpper
            ' Nombre del modulo
            lcHojaExcel.Cells(2, 1).Value = "Ventas"
            ' Titulo del reporte
            lcRango = lcHojaExcel.Range("A3:AD3")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Value = Me.pcNombreReporte
            lcRango.Font.Size = 14
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango.Rows.AutoFit()

            ' Parametros del reporte
            lcRango = lcHojaExcel.Range("A4:AD4")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Value = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
            lcRango.Rows.AutoFit()

            lcRango = lcHojaExcel.Range("K5:O5")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			lcRango = lcHojaExcel.Range("P5:T5")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango = lcHojaExcel.Range("U5:Y5")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango = lcHojaExcel.Range("Z5:AD5")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcHojaExcel.Cells(5, 11).Value = "Por Cuenta de Terceros Multicentro Inversiones, C.A."
            lcHojaExcel.Cells(5, 16).Value = "Por Cuenta de Terceros Condominio El Viñedo, C.A."
            lcHojaExcel.Cells(5, 21).Value = "Por Cuenta de Terceros Multicentro Inversiones, C.A."
            lcHojaExcel.Cells(5, 26).Value = "Perisponio, C.A."

            ' Membrete de la tabla de informacion del reporte
            lcHojaExcel.Cells(6, 2).Value = "Fecha de la Factura"
            lcHojaExcel.Cells(6, 3).Value = "RIF"
            lcHojaExcel.Cells(6, 4).Value = "Nombre o Razón Social"
            lcHojaExcel.Cells(6, 5).Value = "Nº de Factura"
            lcHojaExcel.Cells(6, 6).Value = "Nº de Control"
            lcHojaExcel.Cells(6, 7).Value = "Nº Nota de Débito"
            lcHojaExcel.Cells(6, 8).Value = "Nº Nota de Crédito"
            lcHojaExcel.Cells(6, 9).Value = "Tipo de Transacción"

            lcHojaExcel.Cells(6, 10).Value = "Nº de Factura Afectada"
            lcHojaExcel.Cells(6, 11).Value = "Total Ventas (Incluye IVA)"
            lcHojaExcel.Cells(6, 12).Value = "Ventas Internas No Gravadas"
            lcHojaExcel.Cells(6, 13).Value = "Base Imponible"
            lcHojaExcel.Cells(6, 14).Value = "% Alicuota"
            lcHojaExcel.Cells(6, 15).Value = "Impuesto IVA"

            lcHojaExcel.Cells(6, 16).Value = "Total Ventas (Incluye IVA)"
            lcHojaExcel.Cells(6, 17).Value = "Ventas Internas No Gravadas"
            lcHojaExcel.Cells(6, 18).Value = "Base Imponible"
            lcHojaExcel.Cells(6, 19).Value = "% Alicuota"
            lcHojaExcel.Cells(6, 20).Value = "Impuesto IVA"

            lcHojaExcel.Cells(6, 21).Value = "Total Ventas (Incluye IVA)"
            lcHojaExcel.Cells(6, 22).Value = "Ventas Internas No Gravadas"
            lcHojaExcel.Cells(6, 23).Value = "Base Imponible"
            lcHojaExcel.Cells(6, 24).Value = "% Alicuota"
            lcHojaExcel.Cells(6, 25).Value = "Impuesto IVA"

            lcHojaExcel.Cells(6, 26).Value = "Total Ventas (Incluye IVA)"
            lcHojaExcel.Cells(6, 27).Value = "Ventas Internas No Gravadas"
            lcHojaExcel.Cells(6, 28).Value = "Base Imponible"
            lcHojaExcel.Cells(6, 29).Value = "% Alicuota"
            lcHojaExcel.Cells(6, 30).Value = "Impuesto IVA"

            ' Se le da formato a las celdas del membrete de la tabla
            lcRango = lcHojaExcel.Range("A6:AD6")
            lcRango.Select()
            lcRango.Font.Bold = True

            ' Formato a las columnas de la tabla de informacion del reporte
            lcRango = lcHojaExcel.Range("B1")
            lcRango.EntireColumn.NumberFormat = "DD/MM/YYYY"

            lcRango = lcHojaExcel.Range("C1:J1")
            lcRango.EntireColumn.NumberFormat = "@"

            lcRango = lcHojaExcel.Range("K1:AD1")
            lcRango.EntireColumn.NumberFormat = "###,###,###,##0.00"

            ' Formato a las celdas de la fecha y hora de creacion
            ' Fecha y hora de creacion
            Dim lcFechaCreacion As DateTime = DateTime.Now()
            lcHojaExcel.Cells(1, 30).NumberFormat = "@"
            lcHojaExcel.Cells(1, 30).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcHojaExcel.Cells(1, 30).Value = lcFechaCreacion.ToString("dd/MM/yyyy")
            lcHojaExcel.Cells(2, 30).NumberFormat = "@"
            lcHojaExcel.Cells(2, 30).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcHojaExcel.Cells(2, 30).Value = lcFechaCreacion.ToString("hh:mm:ss tt")

            ' Recorrido de los datos de la consulta a la base de datos para introducir en la tabla del reporte

            ' Numero de la fila de los datos obtenidos de la consulta a la base de datos
            Dim lcFila As Integer
            ' Total de filas de los datos obtenidos de la consulta a la base de datos
            Dim lcTotalFilas As Integer = loDatosReporte.Tables(0).Rows.Count - 1
            ' Datos de la fila de la consulta a la base de datos
            Dim lcDatosFila As DataRow
            ' Control de agrupamiento 1 - (Tipo de Documento)(Tipo_Documento)
            Dim grupo1 As String = ""
            ' Control de agrupamiento 2 - (Codigo del Tipo de Documento)(Cod_Tip)
            Dim grupo2 As String = ""
            ' Numero de la fila en el documento excel
            Dim lcNumFila As Integer = 7
            ' Numero de la fila en el documento excel, inicial para la sumatoria del total para el grupo de control 2
            Dim lcFilaIni As Integer = 0
            ' Numero de la fila en el documento excel, final para la sumatoria del total para el grupo de control 2
            Dim lcFilaFin As Integer = 0
            ' Construccion de la formula de total para el grupo de control 1
            Dim lcTotalDocumento As String = "= 0"

            ' Recorriendo las filas de los datos de la consulta a la base de datos
            For lcFila = 0 To lcTotalFilas - 1
                ' Se extrae los datos de la fila
                lcDatosFila = loDatosReporte.Tables(0).Rows(lcFila)

                ' Si el Tipo_Documento ha cambiado
                If (Trim(lcDatosFila("Tipo_Documento")) <> grupo1) Then
                    ' Si no es la primera vez que ocuerre el cambio
                    If (grupo1 <> "") Then
                        ' Se coloca la etiqueta de total 
                        lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
                        lcRango.Select()
                        lcRango.MergeCells = True
                        lcRango.EntireRow.Font.Bold = True
                        lcRango.Value = "Total del Libro"
                        ' Se coloca la formula de total de todas las columnas
                        ' bloque 1
                        lcHojaExcel.Cells(lcNumFila, 11).Formula = lcTotalDocumento.Replace("[LETRA]", "K")
                        lcHojaExcel.Cells(lcNumFila, 12).Formula = lcTotalDocumento.Replace("[LETRA]", "L")
                        lcHojaExcel.Cells(lcNumFila, 13).Formula = lcTotalDocumento.Replace("[LETRA]", "M")
                        lcHojaExcel.Cells(lcNumFila, 14).Formula = lcTotalDocumento.Replace("[LETRA]", "N")
                        lcHojaExcel.Cells(lcNumFila, 15).Formula = lcTotalDocumento.Replace("[LETRA]", "O")

                        ' bloque 2
                        lcHojaExcel.Cells(lcNumFila, 16).Formula = lcTotalDocumento.Replace("[LETRA]", "P")
                        lcHojaExcel.Cells(lcNumFila, 17).Formula = lcTotalDocumento.Replace("[LETRA]", "Q")
                        lcHojaExcel.Cells(lcNumFila, 18).Formula = lcTotalDocumento.Replace("[LETRA]", "R")
                        lcHojaExcel.Cells(lcNumFila, 19).Formula = lcTotalDocumento.Replace("[LETRA]", "S")
                        lcHojaExcel.Cells(lcNumFila, 20).Formula = lcTotalDocumento.Replace("[LETRA]", "T")

                        ' bloque 3
                        lcHojaExcel.Cells(lcNumFila, 21).Formula = lcTotalDocumento.Replace("[LETRA]", "U")
                        lcHojaExcel.Cells(lcNumFila, 22).Formula = lcTotalDocumento.Replace("[LETRA]", "V")
                        lcHojaExcel.Cells(lcNumFila, 23).Formula = lcTotalDocumento.Replace("[LETRA]", "W")
                        lcHojaExcel.Cells(lcNumFila, 24).Formula = lcTotalDocumento.Replace("[LETRA]", "X")
                        lcHojaExcel.Cells(lcNumFila, 25).Formula = lcTotalDocumento.Replace("[LETRA]", "Y")

                        ' bloque 4
                        lcHojaExcel.Cells(lcNumFila, 26).Formula = lcTotalDocumento.Replace("[LETRA]", "Z")
                        lcHojaExcel.Cells(lcNumFila, 27).Formula = lcTotalDocumento.Replace("[LETRA]", "AA")
                        lcHojaExcel.Cells(lcNumFila, 28).Formula = lcTotalDocumento.Replace("[LETRA]", "AB")
                        lcHojaExcel.Cells(lcNumFila, 29).Formula = lcTotalDocumento.Replace("[LETRA]", "AC")
                        lcHojaExcel.Cells(lcNumFila, 30).Formula = lcTotalDocumento.Replace("[LETRA]", "AD")

                        ' Se inicicializa la variable de control
                        lcTotalDocumento = "= 0"
                        lcNumFila = lcNumFila + 2

                    End If

                    ' Se almacena el dato  en la variable de control para el grupo 1
                    grupo1 = Trim(lcDatosFila("Tipo_Documento"))

                    ' Se coloc el titulo del grupo
                    lcRango = lcHojaExcel.Range("B" & CStr(lcNumFila) & ":D" & CStr(lcNumFila))
                    lcRango.Select()
                    lcRango.MergeCells = True
                    lcRango.EntireRow.Font.Bold = True
                    lcRango.EntireRow.RowHeight = 30
                    lcRango.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    lcRango.EntireRow.NumberFormat = "@"
                    lcHojaExcel.Cells(lcNumFila, 1).Value = grupo1

                    If (grupo1 = "01") Then
                        lcRango.Value = "Retenciones de I.V.A. a Documentos fuera del Período"
                    Else
                        lcRango.Value = "Documentos del Libro de Ventas"
                    End If

                    lcNumFila = lcNumFila + 1

                End If

                ' Si el Cod_tip ha cambiado
                If (Trim(lcDatosFila("Cod_Tip")) <> grupo2) Then
                    ' Si no es la primera vez que ocuerre el cambio
                    If (grupo2 <> "") Then
                        ' Se almacena el numero de la fila final del grupo de datos
                        lcFilaFin = lcNumFila - 1
                        ' Se coloca la etiqueta de total 
                        lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
                        lcRango.Select()
                        lcRango.MergeCells = True
                        lcRango.EntireRow.Font.Bold = True
                        lcRango.Value = "Total del Tipo de Documento"
                        ' Se coloca la formula de total de todas las columnas
                        ' bloque 1
                        lcHojaExcel.Cells(lcNumFila, 11).Formula = "=SUM(K" & CStr(lcFilaIni) & ":K" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 12).Formula = "=SUM(L" & CStr(lcFilaIni) & ":L" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 13).Formula = "=SUM(M" & CStr(lcFilaIni) & ":M" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 14).Formula = "=SUM(N" & CStr(lcFilaIni) & ":N" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 15).Formula = "=SUM(O" & CStr(lcFilaIni) & ":O" & CStr(lcFilaFin) & ")"

                        ' bloque 2
                        lcHojaExcel.Cells(lcNumFila, 16).Formula = "=SUM(P" & CStr(lcFilaIni) & ":P" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 17).Formula = "=SUM(Q" & CStr(lcFilaIni) & ":Q" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 18).Formula = "=SUM(R" & CStr(lcFilaIni) & ":R" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 19).Formula = "=SUM(S" & CStr(lcFilaIni) & ":S" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 20).Formula = "=SUM(T" & CStr(lcFilaIni) & ":T" & CStr(lcFilaFin) & ")"

                        ' bloque 3
                        lcHojaExcel.Cells(lcNumFila, 21).Formula = "=SUM(U" & CStr(lcFilaIni) & ":U" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 22).Formula = "=SUM(V" & CStr(lcFilaIni) & ":V" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 23).Formula = "=SUM(W" & CStr(lcFilaIni) & ":W" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 24).Formula = "=SUM(X" & CStr(lcFilaIni) & ":X" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 25).Formula = "=SUM(Y" & CStr(lcFilaIni) & ":Y" & CStr(lcFilaFin) & ")"

                        ' bloque 4
                        lcHojaExcel.Cells(lcNumFila, 26).Formula = "=SUM(Z" & CStr(lcFilaIni) & ":Z" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 27).Formula = "=SUM(AA" & CStr(lcFilaIni) & ":AA" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 28).Formula = "=SUM(AB" & CStr(lcFilaIni) & ":AB" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 29).Formula = "=SUM(AC" & CStr(lcFilaIni) & ":AC" & CStr(lcFilaFin) & ")"
                        lcHojaExcel.Cells(lcNumFila, 30).Formula = "=SUM(AD" & CStr(lcFilaIni) & ":AD" & CStr(lcFilaFin) & ")"

                        ' Se agrega la celda del total del grupo 2 a la formula del grupo 1
                        lcTotalDocumento = lcTotalDocumento & " + [LETRA]" & CStr(lcNumFila)

                        lcNumFila = lcNumFila + 2

                    End If

                    ' Se almacena el dato  en la variable de control para el grupo 2
                    grupo2 = Trim(lcDatosFila("Cod_Tip"))

                    ' Se coloc el titulo del grupo
                    lcRango = lcHojaExcel.Range("B" & CStr(lcNumFila) & ":D" & CStr(lcNumFila))
                    lcRango.Select()
                    lcRango.MergeCells = True
                    lcRango.EntireRow.Font.Bold = True
                    lcRango.EntireRow.RowHeight = 20
                    lcRango.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    lcRango.EntireRow.NumberFormat = "@"
                    lcRango.Value = "Tipo de Documento: " & grupo2

                    lcNumFila = lcNumFila + 1
                    ' Se almacena el numero de la fila inicial del grupo de datos
                    lcFilaIni = lcNumFila
                End If


                ' Se agrega la informacion en la tabla de los valores detallados
                lcHojaExcel.Cells(lcNumFila, 1).Value = lcDatosFila("Renglon")
                lcHojaExcel.Cells(lcNumFila, 2).Value = lcDatosFila("Fec_Ini")
                lcHojaExcel.Cells(lcNumFila, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                lcHojaExcel.Cells(lcNumFila, 3).Value = lcDatosFila("Rif")
                lcHojaExcel.Cells(lcNumFila, 4).Value = lcDatosFila("Nom_Cli")
                If (Trim(lcDatosFila("Cod_Tip")) = "FACT") Then
                    lcHojaExcel.Cells(lcNumFila, 5).Value = lcDatosFila("Documento")
                Else
                    lcHojaExcel.Cells(lcNumFila, 5).Value = New String("")
                End If
                lcHojaExcel.Cells(lcNumFila, 6).Value = lcDatosFila("Control")
                If (Trim(lcDatosFila("Cod_Tip")) = "N/DB") Then
                    lcHojaExcel.Cells(lcNumFila, 7).Value = lcDatosFila("Documento")
                Else
                    lcHojaExcel.Cells(lcNumFila, 7).Value = New String("")
                End If
                If (Trim(lcDatosFila("Cod_Tip")) = "N/CR") Then
                    lcHojaExcel.Cells(lcNumFila, 8).Value = lcDatosFila("Documento")
                Else
                    lcHojaExcel.Cells(lcNumFila, 8).Value = New String("")
                End If
                lcHojaExcel.Cells(lcNumFila, 9).Value = lcDatosFila("Status_Documento")
                lcHojaExcel.Cells(lcNumFila, 10).Value = lcDatosFila("Doc_Ori")

                ' bloque 1
                If (Trim(lcDatosFila("Cod_Rev")) = "02" And lcDatosFila("Cod_Cla").ToString.Trim().ToLower <> "prop") Then
                    lcHojaExcel.Cells(lcNumFila, 11).Value = lcDatosFila("Mon_Net")
                    lcHojaExcel.Cells(lcNumFila, 12).Value = lcDatosFila("Mon_Exe")
                    lcHojaExcel.Cells(lcNumFila, 13).Value = lcDatosFila("Mon_Bru")
                    lcHojaExcel.Cells(lcNumFila, 14).Value = lcDatosFila("Por_Imp1")
                    lcHojaExcel.Cells(lcNumFila, 15).Value = lcDatosFila("Mon_Imp1")
                Else
                    lcHojaExcel.Cells(lcNumFila, 11).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 12).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 13).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 14).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 15).Value = 0
                End If

                ' bloque 2
                If (Trim(lcDatosFila("Cod_Rev")) = "02" And lcDatosFila("Cod_Cla").ToString.Trim().ToLower = "prop") Then
                    lcHojaExcel.Cells(lcNumFila, 16).Value = lcDatosFila("Mon_Net")
                    lcHojaExcel.Cells(lcNumFila, 17).Value = lcDatosFila("Mon_Exe")
                    lcHojaExcel.Cells(lcNumFila, 18).Value = lcDatosFila("Mon_Bru")
                    lcHojaExcel.Cells(lcNumFila, 19).Value = lcDatosFila("Por_Imp1")
                    lcHojaExcel.Cells(lcNumFila, 20).Value = lcDatosFila("Mon_Imp1")
                Else
                    lcHojaExcel.Cells(lcNumFila, 16).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 17).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 18).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 19).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 20).Value = 0
                End If

                'bloque 3
                If ((Trim(lcDatosFila("Cod_Rev")) = "01" Or Trim(lcDatosFila("Cod_Rev")) = "04") And Not (lcDatosFila("Nom_Cli").ToString().Trim().ToLower() = "cond." Or lcDatosFila("Nom_Cli").ToString.ToLower = "mica")) Then
                    lcHojaExcel.Cells(lcNumFila, 21).Value = lcDatosFila("Mon_Net")
                    lcHojaExcel.Cells(lcNumFila, 22).Value = lcDatosFila("Mon_Exe")
                    lcHojaExcel.Cells(lcNumFila, 23).Value = lcDatosFila("Mon_Bru")
                    lcHojaExcel.Cells(lcNumFila, 24).Value = lcDatosFila("Por_Imp1")
                    lcHojaExcel.Cells(lcNumFila, 25).Value = lcDatosFila("Mon_Imp1")
                Else
                    lcHojaExcel.Cells(lcNumFila, 21).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 22).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 23).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 24).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 25).Value = 0
                End If

                ' bloque 4
                'If ((Trim(lcDatosFila("Cod_Rev")) = "01" Or Trim(lcDatosFila("Cod_Rev")) = "09") And (lcDatosFila("Cod_Cli").ToString().Trim().ToLower() = "cond." Or lcDatosFila("Cod_Cli").ToString.ToLower = "mica")) Then
                If ((Trim(lcDatosFila("Cod_Rev")) = "09") And (lcDatosFila("Cod_Cli").ToString().Trim().ToLower() = "cond." Or lcDatosFila("Cod_Cli").ToString.Trim().ToLower() = "mica")) Then
                    lcHojaExcel.Cells(lcNumFila, 26).Value = lcDatosFila("Mon_Net")
                    lcHojaExcel.Cells(lcNumFila, 27).Value = lcDatosFila("Mon_Exe")
                    lcHojaExcel.Cells(lcNumFila, 28).Value = lcDatosFila("Mon_Bru")
                    lcHojaExcel.Cells(lcNumFila, 29).Value = lcDatosFila("Por_Imp1")
                    lcHojaExcel.Cells(lcNumFila, 30).Value = lcDatosFila("Mon_Imp1")
                Else
                    lcHojaExcel.Cells(lcNumFila, 26).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 27).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 28).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 29).Value = 0
                    lcHojaExcel.Cells(lcNumFila, 30).Value = 0
                End If

                '' bloque 4
                ''If ((Trim(lcDatosFila("Cod_Rev")) = "01" Or Trim(lcDatosFila("Cod_Rev")) = "09") And (lcDatosFila("Cod_Cli").ToString().Trim().ToLower() = "cond." Or lcDatosFila("Cod_Cli").ToString.ToLower = "mica")) Then
                'If ((Trim(lcDatosFila("Cod_Rev")) = "09") And (lcDatosFila("Cod_Cli").ToString().Trim().ToLower() = "cond." Or lcDatosFila("Cod_Cli").ToString.ToLower = "mica")) Then
                '    lcHojaExcel.Cells(lcNumFila, 26).Value = lcDatosFila("Mon_Net")
                '    lcHojaExcel.Cells(lcNumFila, 27).Value = lcDatosFila("Mon_Exe")
                '    lcHojaExcel.Cells(lcNumFila, 28).Value = lcDatosFila("Mon_Bru")
                '    lcHojaExcel.Cells(lcNumFila, 29).Value = lcDatosFila("Por_Imp1")
                '    lcHojaExcel.Cells(lcNumFila, 30).Value = lcDatosFila("Mon_Imp1")
                'Else
                '    lcHojaExcel.Cells(lcNumFila, 26).Value = 0
                '    lcHojaExcel.Cells(lcNumFila, 27).Value = 0
                '    lcHojaExcel.Cells(lcNumFila, 28).Value = 0
                '    lcHojaExcel.Cells(lcNumFila, 29).Value = 0
                '    lcHojaExcel.Cells(lcNumFila, 30).Value = 0
                'End If

                lcNumFila = lcNumFila + 1

            Next lcFila

            ' Se almacena el numero de la fila final del grupo de datos
            lcFilaFin = lcNumFila - 1
            ' Se coloca la etiqueta de total 
            lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.EntireRow.Font.Bold = True
            lcRango.Value = "Total del Tipo de Documento"
            ' Se coloca la formula de total de todas las columnas
            ' bloque 1
            lcHojaExcel.Cells(lcNumFila, 11).Formula = "=SUM(K" & CStr(lcFilaIni) & ":K" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 12).Formula = "=SUM(L" & CStr(lcFilaIni) & ":L" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 13).Formula = "=SUM(M" & CStr(lcFilaIni) & ":M" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 14).Formula = "=SUM(N" & CStr(lcFilaIni) & ":N" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 15).Formula = "=SUM(O" & CStr(lcFilaIni) & ":O" & CStr(lcFilaFin) & ")"

            ' bloque 2
            lcHojaExcel.Cells(lcNumFila, 16).Formula = "=SUM(P" & CStr(lcFilaIni) & ":P" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 17).Formula = "=SUM(Q" & CStr(lcFilaIni) & ":Q" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 18).Formula = "=SUM(R" & CStr(lcFilaIni) & ":R" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 19).Formula = "=SUM(S" & CStr(lcFilaIni) & ":S" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 20).Formula = "=SUM(T" & CStr(lcFilaIni) & ":T" & CStr(lcFilaFin) & ")"

            ' bloque 3
            lcHojaExcel.Cells(lcNumFila, 21).Formula = "=SUM(U" & CStr(lcFilaIni) & ":U" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 22).Formula = "=SUM(V" & CStr(lcFilaIni) & ":V" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 23).Formula = "=SUM(W" & CStr(lcFilaIni) & ":W" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 24).Formula = "=SUM(X" & CStr(lcFilaIni) & ":X" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 25).Formula = "=SUM(Y" & CStr(lcFilaIni) & ":Y" & CStr(lcFilaFin) & ")"

            ' bloque 4
            lcHojaExcel.Cells(lcNumFila, 26).Formula = "=SUM(Z" & CStr(lcFilaIni) & ":Z" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 27).Formula = "=SUM(AA" & CStr(lcFilaIni) & ":AA" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 28).Formula = "=SUM(AB" & CStr(lcFilaIni) & ":AB" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 29).Formula = "=SUM(AC" & CStr(lcFilaIni) & ":AC" & CStr(lcFilaFin) & ")"
            lcHojaExcel.Cells(lcNumFila, 30).Formula = "=SUM(AD" & CStr(lcFilaIni) & ":AD" & CStr(lcFilaFin) & ")"

            ' Se agrega la celda del total del grupo 2 a la formula del grupo 1
            lcTotalDocumento = lcTotalDocumento & " + [LETRA]" & CStr(lcNumFila)

            lcNumFila = lcNumFila + 2

            ' Se coloca la etiqueta de total 
            lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.EntireRow.Font.Bold = True
            lcRango.Value = "Total del Libro"
            ' Se coloca la formula de total de todas las columnas
            ' bloque 1
            lcHojaExcel.Cells(lcNumFila, 11).Formula = lcTotalDocumento.Replace("[LETRA]", "K")
            lcHojaExcel.Cells(lcNumFila, 12).Formula = lcTotalDocumento.Replace("[LETRA]", "L")
            lcHojaExcel.Cells(lcNumFila, 13).Formula = lcTotalDocumento.Replace("[LETRA]", "M")
            lcHojaExcel.Cells(lcNumFila, 14).Formula = lcTotalDocumento.Replace("[LETRA]", "N")
            lcHojaExcel.Cells(lcNumFila, 15).Formula = lcTotalDocumento.Replace("[LETRA]", "O")

            ' bloque 2
            lcHojaExcel.Cells(lcNumFila, 16).Formula = lcTotalDocumento.Replace("[LETRA]", "P")
            lcHojaExcel.Cells(lcNumFila, 17).Formula = lcTotalDocumento.Replace("[LETRA]", "Q")
            lcHojaExcel.Cells(lcNumFila, 18).Formula = lcTotalDocumento.Replace("[LETRA]", "R")
            lcHojaExcel.Cells(lcNumFila, 19).Formula = lcTotalDocumento.Replace("[LETRA]", "S")
            lcHojaExcel.Cells(lcNumFila, 20).Formula = lcTotalDocumento.Replace("[LETRA]", "T")

            ' bloque 3
            lcHojaExcel.Cells(lcNumFila, 21).Formula = lcTotalDocumento.Replace("[LETRA]", "U")
            lcHojaExcel.Cells(lcNumFila, 22).Formula = lcTotalDocumento.Replace("[LETRA]", "V")
            lcHojaExcel.Cells(lcNumFila, 23).Formula = lcTotalDocumento.Replace("[LETRA]", "W")
            lcHojaExcel.Cells(lcNumFila, 24).Formula = lcTotalDocumento.Replace("[LETRA]", "X")
            lcHojaExcel.Cells(lcNumFila, 25).Formula = lcTotalDocumento.Replace("[LETRA]", "Y")

            ' bloque 4
            lcHojaExcel.Cells(lcNumFila, 26).Formula = lcTotalDocumento.Replace("[LETRA]", "Z")
            lcHojaExcel.Cells(lcNumFila, 27).Formula = lcTotalDocumento.Replace("[LETRA]", "AA")
            lcHojaExcel.Cells(lcNumFila, 28).Formula = lcTotalDocumento.Replace("[LETRA]", "AB")
            lcHojaExcel.Cells(lcNumFila, 29).Formula = lcTotalDocumento.Replace("[LETRA]", "AC")
            lcHojaExcel.Cells(lcNumFila, 30).Formula = lcTotalDocumento.Replace("[LETRA]", "AD")

            ' Se agregan los totales del Documento
            'lcNumFila = lcNumFila + 3
            'lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Total Ventas"
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            'lcNumFila = lcNumFila + 1
            'lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Ventas Gravadas"
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            'lcNumFila = lcNumFila + 1
            'lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Total Exentas"
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            'lcNumFila = lcNumFila + 1
            'lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Impuestos"
            'lcHojaExcel.Cells(lcNumFila, 11).Value = "0,00 %"
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            ' Ajustamos el tamaño de las columnas
            lcRango = lcHojaExcel.Range("B1:AD" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.Columns.AutoFit()
			
            ' Seleccionamos la primera celda del libro
            lcRango = lcHojaExcel.Range("A1")
            lcRango.Select()

            ' Cerramos y liberamos recursos
            mLiberar(lcRango)
            mLiberar(lcHojaExcel)
            'Guardamos los cambios del libro activo
            lcLibroExcel.Close(True, loFileName)

            mLiberar(lcLibroExcel)
            loAppExcel.Application.Quit()
            mLiberar(loAppExcel)

        Catch loExcepcion As Exception
                me.mEscribirConsulta(loExcepcion.Message)
                Me.Response.Flush()
                Me.Response.Close()
                
				Me.Response.End()

        Finally
            ' Se forza el cierre del proceso excel
            'Dim Lista_Procesos() As Diagnostics.Process
            'Dim p As Diagnostics.Process
            'Lista_Procesos = Diagnostics.Process.GetProcessesByName("EXCEL")
            'For Each p In Lista_Procesos
            '    Try
            '        p.Kill()
            '    Catch
            '    End Try
            'Next
            GC.Collect()
        End Try

    End Sub


''' <summary>
''' Cierre y liberacion de recursos de los objetos de la libreria Excel
''' </summary>
''' <param name="objeto"></param>
''' <remarks></remarks>
    Private Sub mLiberar(ByVal objeto As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto)
            objeto = Nothing
        Catch ex As Exception
            objeto = Nothing
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
' CMS: 02/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 07/08/10: Se agrego el formato de la salida por excel.
'-------------------------------------------------------------------------------------------'
' RJG: 13/07/11: Corregida error en la salida de excel: hacia referencia a la propiedad		'
'				 goReportes.pcModuloActual (actualmetne eliminada).							'
'-------------------------------------------------------------------------------------------'
