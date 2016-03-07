Imports System.Data
Partial Class rLibro_Ventas_CCMV

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
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" INTO      #tmpLibroVentas ")
            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            lcComandoSeleccionar.AppendLine(" JOIN Renglones_Cobros on (Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip And Cuentas_Cobrar.Documento    = Renglones_Cobros.Doc_Ori) ")
            lcComandoSeleccionar.AppendLine(" JOIN Cobros on Cobros.Documento = Renglones_Cobros.Documento")
            lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Cli = Articulos.Cod_Art ")
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
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN  Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Cli = Articulos.Cod_Art ")
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
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli")
            lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Cli = Articulos.Cod_Art ")
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
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Clientes.Cod_Cli ")

            lcComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            lcComandoSeleccionar.AppendLine(" JOIN Clientes on Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli")
            lcComandoSeleccionar.AppendLine(" LEFT JOIN Articulos On Cuentas_Cobrar.Cod_Cli = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine(" WHERE     ")

            lcComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)

            lcComandoSeleccionar.AppendLine(" SELECT    * ")
            lcComandoSeleccionar.AppendLine(" FROM      #tmpLibroVentas ")
            lcComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

            'lcComandoSeleccionar.AppendLine(" ORDER BY  6, 2 ")


            'Response.Clear()
            'Response.Write("<html><body><pre>" & vbNewLine)
            'Response.Write(lcComandoSeleccionar.ToString)
            'Response.Write("</pre></body></html>")
            'Response.Flush()
            'Response.End()



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            If Me.Request.QueryString("Salida").ToLower = "excel" Then


                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, "C:\Inetpub\wwwroot\eFactory\Administrativo\Complementos\rLibro_Ventas_CCMV.xsl")
                Response.Clear()
                'Response.AddHeader("", "aplication/excel")
                Response.WriteFile("C:\Inetpub\wwwroot\eFactory\Administrativo\Complementos\rLibro_Ventas_CCMV.xsl")
                Response.Flush()
                Response.End()

            End If

            Me.crvrLibro_Ventas_CCMV.ReportSource = loObjetoReporte

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
' CMS: 02/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
