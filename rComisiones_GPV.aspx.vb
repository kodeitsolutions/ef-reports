'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rComisiones_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class rComisiones_GPV

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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("--  Tabulador de Comisiones")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpTabulador(Desde DECIMAL(28,10), Hasta DECIMAL(28,10), Valor DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpTabulador(Desde, Hasta, Valor)")
            loComandoSeleccionar.AppendLine("SELECT  Num_Ini, Num_Fin, Val_Num")
            loComandoSeleccionar.AppendLine("FROM    Renglones_Series")
            loComandoSeleccionar.AppendLine("WHERE   Cod_ser='TABCOMGPV'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("--  Busca las facturas del periodo ")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpComisiones(Posicion            INT IDENTITY(1,1),")
            loComandoSeleccionar.AppendLine("                            Recibo_Caja         CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Cedula              CHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Cod_Cli             CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Desde               DATETIME,")
            loComandoSeleccionar.AppendLine("                            Hasta               DATETIME,")
            loComandoSeleccionar.AppendLine("                            Cliente             VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Factura             CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Monto_Neto          DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            IVA                 DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Dias_Credito        INT,")
            loComandoSeleccionar.AppendLine("                            Descuento           DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Monto_Bruto         DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Descuento1          DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            NCR                 DECIMAL(28,10) DEFAULT(0),")
            loComandoSeleccionar.AppendLine("                            Origen              VARCHAR(50) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loComandoSeleccionar.AppendLine("                            Flete               DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Fraccion_Abonada    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Pago_Efectuado      DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Base_Calculo        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Comision            DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Monto_Comision      DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("                            Fecha_Emision       DATETIME,")
            loComandoSeleccionar.AppendLine("                            Fecha_Vencimiento   DATETIME,")
            loComandoSeleccionar.AppendLine("                            Fecha_Pago          DATETIME,")
            loComandoSeleccionar.AppendLine("                            Cancelacion_En      INT,")
            loComandoSeleccionar.AppendLine("                            Exceso              INT,")
            loComandoSeleccionar.AppendLine("                            Vendedor            CHAR(150) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Cheque_Devuelto     BIT DEFAULT(0),")
            loComandoSeleccionar.AppendLine("                            Con_Comision        BIT DEFAULT(0)")
            loComandoSeleccionar.AppendLine("                            )")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpComisiones( Recibo_Caja, Cedula, Cod_Cli, Desde, Hasta,")
            loComandoSeleccionar.AppendLine("                            Cliente, Factura, Monto_Neto, IVA, ")
            loComandoSeleccionar.AppendLine("                            Dias_Credito, Descuento, Monto_Bruto, Descuento1, ")
            loComandoSeleccionar.AppendLine("                            Flete, Pago_Efectuado, Fraccion_Abonada, Base_Calculo, ")
            loComandoSeleccionar.AppendLine("                            Comision, Monto_Comision, ")
            loComandoSeleccionar.AppendLine("                            Fecha_Emision, Fecha_Vencimiento, Fecha_Pago, Cancelacion_En, Exceso, Vendedor)")
            loComandoSeleccionar.AppendLine("SELECT			Cobros.Documento	                                                            AS Recibo_Caja, ")
            loComandoSeleccionar.AppendLine("				'RIF: ' + Vendedores.cedula                                                     AS Cedula, ")
            loComandoSeleccionar.AppendLine("				Cobros.Cod_Cli                                                                  AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine("				CAST(" & lcParametro0Desde & " AS DATETIME)                                     AS Desde,")
            loComandoSeleccionar.AppendLine("				CAST(" & lcParametro0Hasta & " AS DATETIME)                                     AS Hasta,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli	                                                            AS Cliente, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Documento                                                        AS Factura, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Mon_Bru- Cuentas_Cobrar.Mon_Des                                  AS Monto_Neto, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Mon_Imp1                                                         AS IVA, ")
            loComandoSeleccionar.AppendLine("				DATEDIFF(dd,Cuentas_Cobrar.Fec_Ini,Cuentas_Cobrar.Fec_Fin)                      AS Dias_Credito, ")
            loComandoSeleccionar.AppendLine("				COALESCE(Descuentos_Documentos.Por_Des,0)                                       AS Descuento, ")
            loComandoSeleccionar.AppendLine("				(Cuentas_Cobrar.Mon_Bru- Cuentas_Cobrar.Mon_Des) + Cuentas_Cobrar.Mon_imp1      AS Monto_Bruto, ")
            loComandoSeleccionar.AppendLine("				COALESCE(Descuentos_Documentos.Mon_Des,0)                                       AS Descuento1, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Por_Rec                                                          AS Flete, ")
            loComandoSeleccionar.AppendLine("				Renglones_Cobros.Mon_Abo                                                        AS Pago_Efectuado, ")
            loComandoSeleccionar.AppendLine("				Renglones_Cobros.Mon_Abo/Cuentas_Cobrar.Mon_Net                                 AS Fraccion_Abonada, ")
            loComandoSeleccionar.AppendLine("                ROUND(Renglones_Cobros.Mon_Abo ")
            loComandoSeleccionar.AppendLine("                    - Cuentas_Cobrar.Mon_Imp1*Renglones_Cobros.Mon_Abo/Cuentas_Cobrar.Mon_Net ")
            loComandoSeleccionar.AppendLine("                    - Cuentas_Cobrar.Mon_Rec*Renglones_Cobros.Mon_Abo/Cuentas_Cobrar.Mon_Net ")
            loComandoSeleccionar.AppendLine("                    - COALESCE(Ret_ISLR.Mon_Ret,0) - COALESCE(Ret_Pat.Mon_Ret,0), 2)           AS Base_Calculo,")
            loComandoSeleccionar.AppendLine("				COALESCE(#tmpTabulador.Valor,0)                                                 AS Comision, ")
            loComandoSeleccionar.AppendLine("				( COALESCE(Ret_Pat.Mon_Ret,0) + COALESCE(Ret_ISLR.Mon_Ret,0) ")
            loComandoSeleccionar.AppendLine("                        + COALESCE(Ret_Iva.Mon_Ret,0) + Renglones_Cobros.Mon_Abo ")
            loComandoSeleccionar.AppendLine("                        - Cuentas_Cobrar.Mon_imp1 - Cuentas_Cobrar.Mon_rec")
            loComandoSeleccionar.AppendLine("                    ) * COALESCE(#tmpTabulador.Valor,0) /100                                   AS Monto_Comision, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Fec_Ini                                                          AS Fecha_Emision, ")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.fec_Fin                                                          AS Fecha_Vencimiento, ")
            loComandoSeleccionar.AppendLine("				Cobros.Fec_Ini                                                                  AS Fecha_Pago, ")
            loComandoSeleccionar.AppendLine("				DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cobros.fec_ini)                              AS Cancelacion_En, ")
            loComandoSeleccionar.AppendLine("				DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cobros.fec_ini)")
            loComandoSeleccionar.AppendLine("                    - DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cuentas_Cobrar.fec_fin)               AS Exceso, ")
            loComandoSeleccionar.AppendLine("				'Vendedor: '+ Vendedores.Nom_Ven                                                AS Vendedor ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("FROM			Cobros ")
            loComandoSeleccionar.AppendLine("	JOIN		Vendedores ON Vendedores.cod_ven = Cobros.cod_ven ")
            loComandoSeleccionar.AppendLine("	        AND Cobros.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("	        AND Cobros.Cod_Ven BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("	JOIN		Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli ")
            loComandoSeleccionar.AppendLine("			AND	Cobros.Status = 'Confirmado' ")
            loComandoSeleccionar.AppendLine("	        AND Cobros.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Cobros.Cod_tip = 'FACT' ")
            loComandoSeleccionar.AppendLine("	JOIN		Cuentas_Cobrar ON Cuentas_Cobrar.cod_tip = Renglones_Cobros.cod_tip ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Documento = Renglones_Cobros.doc_ori ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Descuentos_Documentos  ")
            loComandoSeleccionar.AppendLine("			ON  Descuentos_Documentos.Documento= Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Descuentos_Documentos.origen = 'Cobros' ")
            loComandoSeleccionar.AppendLine("			AND Descuentos_Documentos.tipo = Renglones_Cobros.Cod_tip ")
            loComandoSeleccionar.AppendLine("			AND	Descuentos_Documentos.Doc_Ori = Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones_Documentos AS Ret_Iva	 ")
            loComandoSeleccionar.AppendLine("			ON  Ret_Iva.Doc_Ori = Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			AND Ret_Iva.Cla_Ori = Renglones_Cobros.Cod_tip ")
            loComandoSeleccionar.AppendLine("			AND Ret_Iva.Tip_Ori = 'Cuentas_Cobrar' ")
            loComandoSeleccionar.AppendLine("			AND Ret_Iva.Clase = 'Impuesto' ")
            loComandoSeleccionar.AppendLine("			AND Ret_Iva.Documento = Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones_Documentos AS Ret_ISLR	 ")
            loComandoSeleccionar.AppendLine("			ON  Ret_ISLR.Doc_Ori = Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			AND Ret_ISLR.Cla_Ori = Renglones_Cobros.Cod_tip ")
            loComandoSeleccionar.AppendLine("			AND Ret_ISLR.Tip_Ori = 'Cuentas_Cobrar' ")
            loComandoSeleccionar.AppendLine("			AND Ret_ISLR.Clase = 'ISLR' ")
            loComandoSeleccionar.AppendLine("			AND Ret_ISLR.Documento = Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones_Documentos AS Ret_Pat ")
            loComandoSeleccionar.AppendLine("			ON  Ret_Pat.Doc_Ori = Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine("			AND Ret_Pat.Cla_Ori = Renglones_Cobros.Cod_tip ")
            loComandoSeleccionar.AppendLine("			AND Ret_Pat.Tip_Ori = 'Cuentas_Cobrar' ")
            loComandoSeleccionar.AppendLine("			AND Ret_Pat.Clase = 'Patente' ")
            loComandoSeleccionar.AppendLine("			AND Ret_Pat.Documento = Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpTabulador ")
            loComandoSeleccionar.AppendLine("	        ON (CASE ")
            loComandoSeleccionar.AppendLine("					WHEN DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cobros.fec_ini)-DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cuentas_Cobrar.fec_fin) < 0 THEN 0")
            loComandoSeleccionar.AppendLine("					ELSE  DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cobros.fec_ini)-DATEDIFF(dd,Cuentas_Cobrar.fec_ini,Cuentas_Cobrar.fec_fin)")
            loComandoSeleccionar.AppendLine("				END)")
            loComandoSeleccionar.AppendLine("	        BETWEEN #tmpTabulador.Desde AND #tmpTabulador.Hasta")
            loComandoSeleccionar.AppendLine("ORDER BY  Cobros.cod_ven, Vendedores.nom_ven, Cuentas_Cobrar.Fec_Ini")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("--  Resta las NCR a las facturas del mismo cobro")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("UPDATE      #tmpComisiones")
            loComandoSeleccionar.AppendLine("SET         NCR          = C.Base_Calculo_NCR,")
            loComandoSeleccionar.AppendLine("            Origen       = C.Origen,")
            loComandoSeleccionar.AppendLine("            Base_Calculo = Base_Calculo - C.Base_Calculo_NCR")
            loComandoSeleccionar.AppendLine("FROM    (   SELECT      Renglones_Cobros.Documento                                  AS Recibo_Caja,")
            loComandoSeleccionar.AppendLine("                        Renglones_Cobros.Mon_Abo ")
            loComandoSeleccionar.AppendLine("                            - NCR.Mon_Rec*Renglones_Cobros.mon_abo/NCR.mon_net")
            loComandoSeleccionar.AppendLine("                            - NCR.Mon_Imp1*Renglones_Cobros.mon_abo/NCR.mon_net     AS Base_Calculo_NCR,")
            loComandoSeleccionar.AppendLine("                        NCR.Doc_Ori                                                 AS Origen,")
            loComandoSeleccionar.AppendLine("                        Aplicar_A.R                                                 AS Destino")
            loComandoSeleccionar.AppendLine("            FROM        Renglones_Cobros")
            loComandoSeleccionar.AppendLine("                JOIN    #tmpComisiones Creditos ")
            loComandoSeleccionar.AppendLine("                    ON  Creditos.Recibo_Caja = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("                    AND Renglones_Cobros.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                JOIN    Cuentas_Cobrar NCR ")
            loComandoSeleccionar.AppendLine("                    ON  NCR.Documento = Renglones_Cobros.Doc_Ori")
            loComandoSeleccionar.AppendLine("                    AND NCR.Cod_Tip = Renglones_Cobros.Cod_Tip")
            loComandoSeleccionar.AppendLine("                JOIN    (   SELECT      Recibo_Caja, MIN(Posicion) R")
            loComandoSeleccionar.AppendLine("                            FROM        #tmpComisiones")
            loComandoSeleccionar.AppendLine("                            GROUP BY    Recibo_Caja")
            loComandoSeleccionar.AppendLine("                        ) Aplicar_A ")
            loComandoSeleccionar.AppendLine("                    ON  Aplicar_A.Recibo_Caja = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("    ) C")
            loComandoSeleccionar.AppendLine("WHERE   C.Recibo_Caja = #tmpComisiones.Recibo_Caja")
            loComandoSeleccionar.AppendLine("    AND C.Destino = #tmpComisiones.Posicion;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("-- Cheques devueltos con saldo (pendientes)")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("UPDATE  #tmpComisiones")
            loComandoSeleccionar.AppendLine("SET     Cheque_Devuelto = 1,")
            loComandoSeleccionar.AppendLine("        Comision = 0,")
            loComandoSeleccionar.AppendLine("        Monto_Comision = 0,")
            loComandoSeleccionar.AppendLine("        Fecha_Pago = NULL,")
            loComandoSeleccionar.AppendLine("        Cancelacion_En = NULL,")
            loComandoSeleccionar.AppendLine("        Exceso = NULL")
            loComandoSeleccionar.AppendLine("FROM  ( SELECT  #tmpComisiones.Recibo_Caja  AS Recibo_Caja")
            loComandoSeleccionar.AppendLine("        FROM    Cuentas_Cobrar CHEQ")
            loComandoSeleccionar.AppendLine("            JOIN #tmpComisiones ")
            loComandoSeleccionar.AppendLine("            ON  #tmpComisiones.Recibo_Caja = CHEQ.Doc_Ori ")
            loComandoSeleccionar.AppendLine("        WHERE   CHEQ.Cod_Tip = 'CHEQ'")
            loComandoSeleccionar.AppendLine("            AND CHEQ.Tip_Ori = 'Cobros'")
            loComandoSeleccionar.AppendLine("            AND CHEQ.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("            AND CHEQ.Mon_Sal > 0")
            loComandoSeleccionar.AppendLine("        GROUP BY #tmpComisiones.Recibo_Caja")
            loComandoSeleccionar.AppendLine("    ) Devueltos")
            loComandoSeleccionar.AppendLine("WHERE #tmpComisiones.Recibo_Caja = Devueltos.Recibo_Caja;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("-- Cheques devueltos sin saldo (ya pagados)")
            loComandoSeleccionar.AppendLine("-- ********************************************************************")
            loComandoSeleccionar.AppendLine("UPDATE  #tmpComisiones")
            loComandoSeleccionar.AppendLine("SET     Cheque_Devuelto = 1,")
            loComandoSeleccionar.AppendLine("        Comision = COALESCE(#tmpTabulador.Valor, 0),")
            loComandoSeleccionar.AppendLine("        Monto_Comision = ROUND(#tmpComisiones.Base_Calculo*COALESCE(#tmpTabulador.Valor, 0)/100, 2),")
            loComandoSeleccionar.AppendLine("        Fecha_Pago = Cobro_Tardio.Fecha_Cobro,")
            loComandoSeleccionar.AppendLine("        Cancelacion_En = DATEDIFF(dd,#tmpComisiones.Fecha_Emision, Cobro_Tardio.Fecha_Cobro),")
            loComandoSeleccionar.AppendLine("        Exceso = DATEDIFF(dd,#tmpComisiones.Fecha_Emision, Cobro_Tardio.Fecha_Cobro)")
            loComandoSeleccionar.AppendLine("                 - DATEDIFF(dd,#tmpComisiones.Fecha_Emision, #tmpComisiones.Fecha_Vencimiento)")
            loComandoSeleccionar.AppendLine("FROM  ( SELECT      Origen.Recibo_Caja          AS Cobro_Origen,")
            loComandoSeleccionar.AppendLine("                    Origen.Fecha_Emision        AS Fecha_Emision,")
            loComandoSeleccionar.AppendLine("                    Origen.Fecha_Vencimiento    AS Fecha_Vencimiento,")
            loComandoSeleccionar.AppendLine("                    MAX(Cobros.Fec_Ini)         AS Fecha_Cobro")
            loComandoSeleccionar.AppendLine("        FROM        Cuentas_Cobrar CHEQ")
            loComandoSeleccionar.AppendLine("            JOIN    #tmpComisiones Origen")
            loComandoSeleccionar.AppendLine("                ON  Origen.Recibo_Caja = CHEQ.Doc_Ori ")
            loComandoSeleccionar.AppendLine("                AND CHEQ.Cod_Tip = 'CHEQ'")
            loComandoSeleccionar.AppendLine("                AND CHEQ.Tip_Ori = 'Cobros'")
            loComandoSeleccionar.AppendLine("                AND CHEQ.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("                AND CHEQ.Mon_Sal = 0")
            loComandoSeleccionar.AppendLine("            JOIN    Renglones_Cobros ")
            loComandoSeleccionar.AppendLine("                ON  Renglones_Cobros.Cod_Tip = 'CHEQ' ")
            loComandoSeleccionar.AppendLine("                AND Renglones_Cobros.Doc_Ori = CHEQ.Documento")
            loComandoSeleccionar.AppendLine("            JOIN    Cobros ")
            loComandoSeleccionar.AppendLine("                ON  Renglones_Cobros.Documento = Cobros.Documento ")
            loComandoSeleccionar.AppendLine("                AND Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        GROUP BY    Origen.Recibo_Caja,")
            loComandoSeleccionar.AppendLine("                    Origen.Fecha_Emision,")
            loComandoSeleccionar.AppendLine("                    Origen.Fecha_Vencimiento")
            loComandoSeleccionar.AppendLine("    ) Cobro_Tardio")
            loComandoSeleccionar.AppendLine("    LEFT JOIN #tmpTabulador ")
            loComandoSeleccionar.AppendLine("    ON (CASE ")
            loComandoSeleccionar.AppendLine("	    WHEN DATEDIFF(dd,Cobro_Tardio.Fecha_Emision, Cobro_Tardio.Fecha_Cobro)-DATEDIFF(dd,Cobro_Tardio.Fecha_Emision, Cobro_Tardio.Fecha_Vencimiento) < 0 THEN 0")
            loComandoSeleccionar.AppendLine("		ELSE  DATEDIFF(dd,Cobro_Tardio.Fecha_Emision, Cobro_Tardio.Fecha_Cobro)-DATEDIFF(dd,Cobro_Tardio.Fecha_Emision, Cobro_Tardio.Fecha_Vencimiento)")
            loComandoSeleccionar.AppendLine("	END) BETWEEN #tmpTabulador.Desde And #tmpTabulador.Hasta")
            loComandoSeleccionar.AppendLine("WHERE #tmpComisiones.Recibo_Caja = Cobro_Tardio.Cobro_Origen")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE  #tmpComisiones")
            loComandoSeleccionar.AppendLine("SET     Con_Comision = (CASE WHEN Comision > 0 THEN 1 ELSE 0 END);")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpComisiones")
            loComandoSeleccionar.AppendLine("ORDER BY Con_Comision DESC, Posicion")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComisiones_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComisiones_GPV.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' EAG:  18/08/15: Código inicial.								                            '
'-------------------------------------------------------------------------------------------'
' EAG:  22/09/15: Se agregaron los descuentos por N/CR (DxPP).								'
'-------------------------------------------------------------------------------------------'
' RJG:  22/09/15: Continuación de la programación: aplicación de N/CR de Devoluciones, y    '
'                 cheques devueltos.								                        '
'-------------------------------------------------------------------------------------------'
' RJG:  23/09/15: Continuación de la programación: aplicación de N/CR de Devoluciones, y    '
'                 cheques devueltos.								                        '
'-------------------------------------------------------------------------------------------'
