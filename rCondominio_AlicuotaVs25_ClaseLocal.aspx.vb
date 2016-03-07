'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCondominio_AlicuotaVs25_ClaseLocal"
'-------------------------------------------------------------------------------------------'
Partial Class rCondominio_AlicuotaVs25_ClaseLocal
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		Try
	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            
            Dim ldFechaInicial AS Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
    
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lnMontoArrendConReserva DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoArrendSinReserva DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoGastosComunes DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoPropietarios DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoMinitienda DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoOpcionVenta DECIMAL(28,10);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnMontoArrendConReserva = " & lcParametro4Desde & ";")  'Cod_Cla = T1CO
            loConsulta.AppendLine("SET @lnMontoArrendSinReserva = " & lcParametro5Desde & ";")  'Cod_Cla = T1GC
            loConsulta.AppendLine("SET @lnMontoGastosComunes = " & lcParametro6Desde & ";")     'Cod_Cla = T3
            loConsulta.AppendLine("SET @lnMontoPropietarios = " & lcParametro7Desde & ";")      'Cod_Cla = PROP
            loConsulta.AppendLine("SET @lnMontoMinitienda = " & lcParametro8Desde & ";")        '* Modelo = T3
            loConsulta.AppendLine("SET @lnMontoOpcionVenta = " & lcParametro9Desde & ";")       'Cod_Cla = OPVE
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lcMes VARCHAR(2);")
            loConsulta.AppendLine("DECLARE @lcAnio VARCHAR(4);")
            loConsulta.AppendLine("SET @lcMes = '" & ldFechaInicial.Month.ToString() & "';")
            loConsulta.AppendLine("SET @lcAnio = '" & ldFechaInicial.Year.ToString() & "';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  CxC.documento               AS documento,")
            loConsulta.AppendLine("        CxC.Cod_Cli                 AS Cod_Cli,")
            loConsulta.AppendLine("        C.nom_cli                   AS nom_cli,")
            loConsulta.AppendLine("        CxC.mon_bru                 AS mon_bru,")
            loConsulta.AppendLine("        CxC.mon_imp1                AS mon_imp1,")
            loConsulta.AppendLine("        CxC.mon_net                 AS mon_net,")
            loConsulta.AppendLine("        A.Cod_Cla                   AS Cod_Cla,")
            loConsulta.AppendLine("        CA.nom_cla                  AS nom_cla,")
            loConsulta.AppendLine("        A.Modelo                    AS Modelo,")
            loConsulta.AppendLine("        CxC.comentario              AS comentario,")
            loConsulta.AppendLine("        A.Cod_Art                   AS Local,")
            loConsulta.AppendLine("        A.Precio1                   AS Canon,")
            loConsulta.AppendLine("        ROUND(A.precio1*.25,2)      AS Canon_25,")
            loConsulta.AppendLine("        A.Por_Ali                   AS Alicuota1, ")
            loConsulta.AppendLine("        A.Por_Ali2                  AS Alicuota2,")
            loConsulta.AppendLine("        ROUND(A.por_ali*(")
            loConsulta.AppendLine("            CASE A.Cod_Cla ")
            loConsulta.AppendLine("            WHEN 'T1CO' THEN @lnMontoArrendConReserva ")
            loConsulta.AppendLine("            WHEN 'T1GC' THEN @lnMontoArrendSinReserva ")
            loConsulta.AppendLine("            WHEN 'T3'   THEN @lnMontoGastosComunes ")
            loConsulta.AppendLine("            WHEN 'PROP' THEN @lnMontoPropietarios ")
            loConsulta.AppendLine("            --WHEN 'T3'   THEN @lnMontoMinitienda")
            loConsulta.AppendLine("            WHEN 'OPVE' THEN @lnMontoOpcionVenta ")
            loConsulta.AppendLine("            ELSE 0 END)/100,2)      AS Condominio_A1,")
            loConsulta.AppendLine("        ROUND(A.por_ali2*(")
            loConsulta.AppendLine("            CASE WHEN A.Modelo ='T3' OR A.Cod_Cla = 'T3'")
            loConsulta.AppendLine("                THEN @lnMontoMinitienda ")
            loConsulta.AppendLine("                ELSE 0 ")
            loConsulta.AppendLine("            END)/100,2)             AS Condominio_A2,")
            loConsulta.AppendLine("            CxC.fec_ini,")
            loConsulta.AppendLine("            (CASE WHEN CxC.Cod_Cli = A.Cod_Art")
            loConsulta.AppendLine("                 THEN 1 ")
            loConsulta.AppendLine("                 ELSE 0 ")
            loConsulta.AppendLine("            END )                    AS Principal,")
            loConsulta.AppendLine("            CxC.caracter1, CxC.caracter2, CxC.caracter3")
            loConsulta.AppendLine("INTO    #tmpDocumentos")
            loConsulta.AppendLine("FROM    cuentas_cobrar AS CxC")
            loConsulta.AppendLine("    JOIN clientes AS C ON C.Cod_Cli = CxC.Cod_Cli")
            loConsulta.AppendLine("    JOIN precios_clientes PC ON PC.cod_reg = CxC.Cod_Cli AND PC.origen = 'clientes'")
            loConsulta.AppendLine("    JOIN articulos A ON A.cod_art = PC.Cod_Art ")
            loConsulta.AppendLine("    JOIN clases_articulos CA ON CA.Cod_Cla = A.Cod_Cla ")
            loConsulta.AppendLine("WHERE   CxC.caracter3 = 'CONDOMINIO'")
            loConsulta.AppendLine("    AND CxC.caracter1 = @lcMes ")
            loConsulta.AppendLine("    AND CxC.caracter2 = @lcAnio ")
            loConsulta.AppendLine("    AND CxC.Status <> 'Anulado'")
            loConsulta.AppendLine("    AND CxC.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND CxC.cod_cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND A.Cod_Cla BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("ORDER BY A.Cod_Cla ASC,  CxC.documento ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      #tmpDocumentos.Documento,")
            loConsulta.AppendLine("            CAST(CASE WHEN (Primeros_Documentos.Documento IS NULL) ")
            loConsulta.AppendLine("                THEN 0 ELSE 1 END AS BIT) AS Primer_Documento, ")
            loConsulta.AppendLine("            #tmpDocumentos.Cod_Cli, ")
            loConsulta.AppendLine("            #tmpDocumentos.Nom_cli,")
            loConsulta.AppendLine("            #tmpDocumentos.Mon_Bru,")
            loConsulta.AppendLine("            #tmpDocumentos.Mon_Imp1,")
            loConsulta.AppendLine("            (CASE WHEN #tmpDocumentos.Principal = 1")
            loConsulta.AppendLine("                 THEN (#tmpDocumentos.Mon_Net + COALESCE(CxC_Manuales.Monto_Manual,0)) ")
            loConsulta.AppendLine("                 ELSE 0 END) AS Mon_Net,")
            'loConsulta.AppendLine("            (#tmpDocumentos.Mon_Net + COALESCE(CxC_Manuales.Monto_Manual,0)) AS Mon_Net,")
            loConsulta.AppendLine("            #tmpDocumentos.Cod_Cla,")
            loConsulta.AppendLine("            #tmpDocumentos.Nom_Cla,")
            loConsulta.AppendLine("            #tmpDocumentos.Modelo,")
            loConsulta.AppendLine("            #tmpDocumentos.Comentario,")
            loConsulta.AppendLine("            #tmpDocumentos.Local,")
            loConsulta.AppendLine("            (CASE WHEN #tmpDocumentos.Principal = 1")
            loConsulta.AppendLine("                 THEN #tmpDocumentos.Canon ELSE 0 END) AS Canon,")
            loConsulta.AppendLine("            (CASE WHEN #tmpDocumentos.Principal = 1")
            loConsulta.AppendLine("                 THEN #tmpDocumentos.Canon_25 ELSE 0 END) AS Canon_25,")
            loConsulta.AppendLine("            COALESCE(Locales.Num_Locales, 0) AS Num_Locales,")
            loConsulta.AppendLine("            #tmpDocumentos.Alicuota1,")
            loConsulta.AppendLine("            #tmpDocumentos.Alicuota2,")
            loConsulta.AppendLine("            #tmpDocumentos.Condominio_A1,")
            loConsulta.AppendLine("            #tmpDocumentos.Condominio_A2,")
            loConsulta.AppendLine("            (CASE WHEN #tmpDocumentos.Principal = 1")
            loConsulta.AppendLine("                 THEN COALESCE(Locales.Condominio_T, 0) ELSE 0 END) AS Condominio_T,")
            loConsulta.AppendLine("            #tmpDocumentos.Caracter1                             AS Mes,")
            loConsulta.AppendLine("            #tmpDocumentos.Caracter2                             AS Anio,")
            loConsulta.AppendLine("            #tmpDocumentos.Caracter3                             AS Condominio,")
            loConsulta.AppendLine("            (CASE #tmpDocumentos.Cod_Cla ")
            loConsulta.AppendLine("                 WHEN 'T1CO' THEN @lnMontoArrendConReserva")
            loConsulta.AppendLine("                 WHEN 'T1GC' THEN @lnMontoArrendSinReserva")
            loConsulta.AppendLine("                 WHEN 'T3'   THEN @lnMontoGastosComunes ")
            loConsulta.AppendLine("                 WHEN 'PROP' THEN @lnMontoPropietarios ")
            loConsulta.AppendLine("               --WHEN 'T3'   THEN @lnMontoMinitienda")
            loConsulta.AppendLine("                 WHEN 'OPVE' THEN @lnMontoOpcionVenta ")
            loConsulta.AppendLine("            ELSE 0 END)                                          AS Monto_Aplicado1,")
            loConsulta.AppendLine("            @lnMontoMinitienda                                   AS Monto_Aplicado2")
            loConsulta.AppendLine("FROM        #tmpDocumentos")
            loConsulta.AppendLine("    LEFT JOIN(  SELECT  COUNT(*) AS Num_Locales,")
            loConsulta.AppendLine("                        L.Documento, ")
            loConsulta.AppendLine("                        L.Cod_Cli, ")
            loConsulta.AppendLine("                        SUM(L.Condominio_A1)")
            loConsulta.AppendLine("                        + SUM(L.Condominio_A2) AS Condominio_T ")
            loConsulta.AppendLine("                FROM    #tmpDocumentos L")
            loConsulta.AppendLine("                --WHERE   L.Principal = 1")
            loConsulta.AppendLine("                GROUP BY L.Documento, L.Cod_Cli")
            loConsulta.AppendLine("            ) AS Locales")
            loConsulta.AppendLine("        ON  Locales.Documento = #tmpDocumentos.Documento")
            loConsulta.AppendLine("        AND Locales.Cod_Cli = #tmpDocumentos.Local")
            loConsulta.AppendLine("    LEFT JOIN(  SELECT  MIN(L.Documento) AS Documento,")
            loConsulta.AppendLine("                        L.Cod_Cli  ")
            loConsulta.AppendLine("                FROM    #tmpDocumentos L")
            loConsulta.AppendLine("                GROUP BY L.Cod_Cli")
            loConsulta.AppendLine("            ) AS Primeros_Documentos")
            loConsulta.AppendLine("        ON  Primeros_Documentos.Documento = #tmpDocumentos.Documento ")
            loConsulta.AppendLine("        AND Primeros_Documentos.Cod_Cli = #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine("    LEFT JOIN (  SELECT  CxC_Manual.Cod_Cli, ")
            loConsulta.AppendLine("                     SUM(CxC_Manual.Mon_Net) AS Monto_Manual ")
            loConsulta.AppendLine("             FROM    Cuentas_Cobrar CxC_Manual")
            loConsulta.AppendLine("             JOIN    #tmpDocumentos Doc_Condominio")
            loConsulta.AppendLine("                 ON  Doc_Condominio.Cod_Cli = CxC_Manual.Cod_Cli")
            loConsulta.AppendLine("             WHERE   CxC_Manual.Comentario LIKE '%' + RTRIM(Doc_Condominio.Caracter2) + '%'")
            loConsulta.AppendLine("                 AND CxC_Manual.Comentario LIKE '%' + ( CASE RTRIM(Doc_Condominio.Caracter1)")
            loConsulta.AppendLine("                                                             WHEN '1' THEN 'ENERO' ")
            loConsulta.AppendLine("                                                             WHEN '2' THEN 'FEBRERO' ")
            loConsulta.AppendLine("                                                             WHEN '3' THEN 'MARZO' ")
            loConsulta.AppendLine("                                                             WHEN '4' THEN 'ABRIL' ")
            loConsulta.AppendLine("                                                             WHEN '5' THEN 'MAYO' ")
            loConsulta.AppendLine("                                                             WHEN '6' THEN 'JUNIO' ")
            loConsulta.AppendLine("                                                             WHEN '7' THEN 'JULIO' ")
            loConsulta.AppendLine("                                                             WHEN '8' THEN 'AGOSTO' ")
            loConsulta.AppendLine("                                                             WHEN '9' THEN 'SEPTIEMBRE' ")
            loConsulta.AppendLine("                                                             WHEN '10' THEN 'OCTUBRE' ")
            loConsulta.AppendLine("                                                             WHEN '11' THEN 'NOVIEMBRE' ")
            loConsulta.AppendLine("                                                             WHEN '12' THEN 'DICIEMBRE' ")
            loConsulta.AppendLine("                                                        END) + '%'")
            loConsulta.AppendLine("                 AND CxC_Manual.Caracter1 = ''")
            loConsulta.AppendLine("                 AND CxC_Manual.Cod_Tip = 'FACT'")
            loConsulta.AppendLine("                 AND 1=0")
            loConsulta.AppendLine("             GROUP BY CxC_Manual.Cod_Cli")
            loConsulta.AppendLine("             ")
            loConsulta.AppendLine("     ) AS CxC_Manuales")
            loConsulta.AppendLine("        ON  CxC_Manuales.Cod_Cli = #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine("ORDER BY    Cod_Cla, Documento, Num_Locales DESC, Local")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--DROP TABLE #tmpDocumentos")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("") 

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCondominio_AlicuotaVs25_ClaseLocal", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrCondominio_AlicuotaVs25_ClaseLocal.ReportSource = loObjetoReporte	


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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 04/04/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 10/04/14: Ajustes varios: presentación de detalle y resumen, ajustes en totales.     '
'-------------------------------------------------------------------------------------------'
