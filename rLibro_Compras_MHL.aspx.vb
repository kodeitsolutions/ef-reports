'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Compras_MHL"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Compras_MHL

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosFinales(6)
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Declaracion e inicializacion de variables y creacion de tabla temporal de trabajo
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10);")
            loComandoSeleccionar.AppendLine("SET @lnCero = CAST(0 AS DECIMAL(28, 10));")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio AS NVARCHAR(30);")
            loComandoSeleccionar.AppendLine("SET @lcVacio = N'';")

            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpLibroCompras(	Tabla	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Tip 	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Codigo_Tipo	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Documento	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Control	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Referencia	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Factura 	        VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Status		        VARCHAR(15) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Doc_Ori		        VARCHAR(50) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Documento_Afectado	VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Exp_Importacion	    VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Pro		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ini		        DATETIME,")
            loComandoSeleccionar.AppendLine("								Tip_Doc		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Mon_Bru 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Net 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Dis_Imp 	        VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("								Mon_Des 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Rec 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Des 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Rec 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Cod_Imp1 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Imp2 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Imp3 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Com_Ret		    	VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ret		    	DATETIME,")
            loComandoSeleccionar.AppendLine("								mon_exe 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_ret 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_exe 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_bas 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_imp 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Exonerado 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								No_Sujeto 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Sin_Derecho_CF      DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Comentario 	        VARCHAR(250) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Nom_Pro 	        VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Rif 	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Tip_Pro 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Per 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Transaccion		    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Impuesto_Excluir	DECIMAL(28,10) DEFAULT(0))")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpLibroCompras(	Tabla, Cod_Tip, Codigo_Tipo, Documento, Control, Referencia, Factura, Status,")
            loComandoSeleccionar.AppendLine("								Doc_Ori, Documento_Afectado, Exp_Importacion, Cod_Pro, Fec_Ini, Tip_Doc, Mon_Bru, Mon_Net,")
            loComandoSeleccionar.AppendLine("								Dis_Imp, Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loComandoSeleccionar.AppendLine("								Cod_Imp1, Cod_Imp2, Cod_Imp3, mon_exe, Com_Ret, mon_ret, mon_imp1, mon_bas1, por_imp1,")
            loComandoSeleccionar.AppendLine("								mon_exe1, mon_imp2, mon_bas2, por_imp2, mon_exe2, mon_imp3, mon_bas3,")
            loComandoSeleccionar.AppendLine("								por_imp3, mon_exe3, subt_exe, subt_bas, subt_imp, Exonerado, No_Sujeto,")
            loComandoSeleccionar.AppendLine("								Sin_Derecho_CF, Comentario, Nom_Pro, Rif, Tip_Pro, Cod_Per, Transaccion)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("SELECT		")
            loComandoSeleccionar.AppendLine("			'Compras'															            AS Tabla, 	        	")
            loComandoSeleccionar.AppendLine("			CASE Cuentas_Pagar.Cod_Tip														                    	")
            loComandoSeleccionar.AppendLine("			 	WHEN 'FACT' 	THEN 'Factura'													                    ")
            loComandoSeleccionar.AppendLine("			 	WHEN 'N/CR' 	THEN 'Nota de Credito'          										        	")
            loComandoSeleccionar.AppendLine("			 	WHEN 'N/DB' 	THEN 'Nota de Debito'			            							        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Cod_Tip,	        	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tip												            AS Codigo_Tipo,         ")
            loComandoSeleccionar.AppendLine("			CAST(Cuentas_Pagar.Documento AS CHAR(30))							            AS Documento,        	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control												            AS Control,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Referencia											            AS Referencia,	        ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura												            AS Factura,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Status												            AS Status, 	        	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Ori												            AS Doc_Ori,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori												            AS Documento_Afectado, 	")
            loComandoSeleccionar.AppendLine("			@lcVacio															            AS Exp_Importacion, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro												            AS Cod_Pro,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini												            AS Fec_Ini,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc												            AS Tip_Doc,         	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 					            				        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Bru * -1 									            		        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Bru 														                    ")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Bru,          	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 	            								        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net * -1              											        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Net 								            					        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Net,          	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Dis_Imp												            AS Dis_Imp,         	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito'              									        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Des * -1 					            						        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Des 										            			        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Des,          	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Rec												            AS Mon_Rec,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Des												            AS Por_Des,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Rec												            AS Por_Rec,         	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito'              									        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr1 * -1 					            						        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr1 									            			        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Otr1,          	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito'              									        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr2 * -1 					            						        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr2 									            			        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Otr2,          	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito'              									        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr3 * -1 					            						        	")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr3 									            			        	")
            loComandoSeleccionar.AppendLine("			END																	            AS Mon_Otr3,          	")
            loComandoSeleccionar.AppendLine("			@lcVacio															            AS Cod_Imp1,         	")
            loComandoSeleccionar.AppendLine("			@lcVacio															            AS Cod_Imp2,         	")
            loComandoSeleccionar.AppendLine("			@lcVacio															            AS Cod_Imp3,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Exe												            AS Mon_Exe,         	")
            loComandoSeleccionar.AppendLine("			@lcVacio															            AS Com_Ret,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS mon_ret,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1												            AS mon_imp1,         	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Mon_Imp1 > 0.00                  									        	")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Bru      			        									            ")
            loComandoSeleccionar.AppendLine("				ELSE 0.00                        					        							        	")
            loComandoSeleccionar.AppendLine("			END	* (CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)          AS mon_bas1,         	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Imp1 														    AS por_imp1,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS mon_exe1,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS mon_imp2,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS mon_bas2,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS por_imp2,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS mon_exe2,         	")
            loComandoSeleccionar.AppendLine("           @lnCero 															            AS mon_imp3,        	")
            loComandoSeleccionar.AppendLine("           @lnCero 															            AS mon_bas3,        	")
            loComandoSeleccionar.AppendLine("           @lnCero 															            AS por_imp3,        	")
            loComandoSeleccionar.AppendLine("           @lnCero 															            AS mon_exe3,        	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS subt_exe,        	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS subt_bas,        	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS subt_imp,        	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS Exonerado,         	")
            loComandoSeleccionar.AppendLine("			@lnCero 															            AS No_Sujeto,        	")
            loComandoSeleccionar.AppendLine("			@lnCero                                                                         AS Sin_Derecho_CF,	    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Comentario                                                        AS Comentario,          ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 			                    	")
            loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro 													                    	")
            loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') 								                    	")
            loComandoSeleccionar.AppendLine("					THEN Proveedores.Nom_Pro 												                    	")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Nom_Pro 												                    	")
            loComandoSeleccionar.AppendLine("				END) 																		                    	")
            loComandoSeleccionar.AppendLine("			END)																            AS Nom_Pro,             ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 			                    	")
            loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif ELSE 													                    	")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Cuentas_Pagar.Rif = '')											                    	")
            loComandoSeleccionar.AppendLine("					THEN Proveedores.Rif 													                    	")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Rif 													                    	")
            loComandoSeleccionar.AppendLine("			    END) 																		                    	")
            loComandoSeleccionar.AppendLine("			 END)																            AS Rif,		            ")
            loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pro 												            AS Tip_Pro,	            ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '1'									                                ")
            loComandoSeleccionar.AppendLine("				THEN 'NR' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE  																                                ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '2'										                            ")
            loComandoSeleccionar.AppendLine("				THEN 'NNR' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE  																                                ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '3'										                            ")
            loComandoSeleccionar.AppendLine("				THEN 'JD' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE  																                                ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '4'										                            ")
            loComandoSeleccionar.AppendLine("				THEN 'JND' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE  																                                ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '5'										                            ")
            loComandoSeleccionar.AppendLine("				THEN 'EXE' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE  																                                ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '6'										                            ")
            loComandoSeleccionar.AppendLine("				THEN 'TN' 																                            ")
            loComandoSeleccionar.AppendLine("				ELSE 'OTR' 																                            ")
            loComandoSeleccionar.AppendLine("			END)															                                        ")
            loComandoSeleccionar.AppendLine("			END)															                                        ")
            loComandoSeleccionar.AppendLine("			END)															                                        ")
            loComandoSeleccionar.AppendLine("			END)															                                        ")
            loComandoSeleccionar.AppendLine("			END)															                                        ")
            loComandoSeleccionar.AppendLine("			END)															                AS Cod_Per,	            ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Pagar.Status = 'Anulado' 										                    ")
            loComandoSeleccionar.AppendLine("				THEN '03-ANU' 																	                    ")
            loComandoSeleccionar.AppendLine("				ELSE '01-REG' 																                    	")
            loComandoSeleccionar.AppendLine("			END)		            														AS Transaccion	        ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Fec_Reg       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_pro   BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Suc   BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status    IN ( " & lcParametro8Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev   BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev   NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND (										")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/DB')")
            loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 1 AND Cuentas_Pagar.Tip_Ori = 'Devoluciones_Proveedores')")
            loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 0 AND Cuentas_Pagar.Cod_Rev IN ('DEVCOM', 'REBAJA') )")
            loComandoSeleccionar.AppendLine("				)										")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_tip IN ('FACT', 'N/CR', 'N/DB')")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Tip      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim lcStatus As String = "Pendiente,Confirmado,Procesado,Pagado,Cerrado,Afectado,Serializado,Contabilizado,Iniciado,Conciliado,Otro,Anulado"

            If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus)) OrElse _
             (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then

                loComandoSeleccionar.AppendLine("UNION ALL")

                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''' Obtencion de las Facturas de Compra Anuladas ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                loComandoSeleccionar.AppendLine("SELECT		")
                loComandoSeleccionar.AppendLine("			'Compras'															AS Tabla, 		")
                loComandoSeleccionar.AppendLine("			CASE Cuentas_Pagar.Cod_Tip															")
                loComandoSeleccionar.AppendLine("			 	WHEN 'FACT' 	THEN 'Factura'													")
                loComandoSeleccionar.AppendLine("			 	WHEN 'N/CR' 	THEN 'Nota de Credito'											")
                loComandoSeleccionar.AppendLine("			 	WHEN 'N/DB' 	THEN 'Nota de Debito'											")
                loComandoSeleccionar.AppendLine("			END																	AS Cod_Tip,		")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tip												AS Codigo_Tipo,	")
                loComandoSeleccionar.AppendLine("			CAST(Cuentas_Pagar.Documento AS CHAR(30))							AS Documento, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control												AS Control, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Referencia											AS Referencia,	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura												AS Factura, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Status												AS Status, 		")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Ori												AS Doc_Ori,         	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori												AS Documento_Afectado, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio															AS Exp_Importacion, 	")
                'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori												AS Documento_Afectado, 	")
                'loComandoSeleccionar.AppendLine("			@lcVacio															AS Exp_Importacion, 	")
                'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro												AS Cod_Pro, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro												AS Cod_Pro, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini												AS Fec_Ini, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc												AS Tip_Doc, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Bru,  	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Net,  	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Dis_Imp												AS Dis_Imp, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Des,  	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Rec, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Des												AS Por_Des, 	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Rec												AS Por_Rec, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Otr1,  	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Otr2,  	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Otr3,  	")
                loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp1, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp2, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp3, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Mon_Exe, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio															AS Com_Ret, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_ret, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_imp1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_bas1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS por_imp1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_exe1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_imp2, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_bas2, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS por_imp2, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_exe2, 	")
                loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_imp3,	")
                loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_bas3,	")
                loComandoSeleccionar.AppendLine("           @lnCero 															AS por_imp3,	")
                loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_exe3,	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_exe,	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_bas,	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_imp,	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS Exonerado, 	")
                loComandoSeleccionar.AppendLine("			@lnCero 															AS No_Sujeto, 	")
                loComandoSeleccionar.AppendLine("			@lnCero                                                             AS Sin_Derecho_CF,	")
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Comentario                                            AS Comentario, ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro 														")
                loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') 									")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Nom_Pro 													")
                loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Nom_Pro 													")
                loComandoSeleccionar.AppendLine("				END) 																			")
                loComandoSeleccionar.AppendLine("			END)																AS Nom_Pro, 	")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif ELSE 														")
                loComandoSeleccionar.AppendLine("			    (CASE WHEN (Cuentas_Pagar.Rif = '')												")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Rif 														")
                loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Rif 														")
                loComandoSeleccionar.AppendLine("			    END) 																			")
                loComandoSeleccionar.AppendLine("			 END)																AS Rif,			")
                loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pro 												AS Tip_Pro,		")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '1'									        ")
                loComandoSeleccionar.AppendLine("				THEN 'NR' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '2'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'NNR' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '3'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'JD' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '4'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'JND' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '5'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'EXE' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '6'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'TN' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE 'OTR' 																    ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															    AS Cod_Per,	")
                loComandoSeleccionar.AppendLine("			'03-ANU'    														AS Transaccion	")
                loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar ")
                loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
                loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Fec_Reg BETWEEN " & lcParametro0Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
                loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status        IN ( " & lcParametro8Desde & " ) ")
                If lcParametro5Desde = "Igual" Then
                    loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
                loComandoSeleccionar.AppendLine("			AND (										")
                loComandoSeleccionar.AppendLine("						Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/DB')")
                loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 1 AND Cuentas_Pagar.Tip_Ori = 'Devoluciones_Proveedores')")
                loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 0 AND Cuentas_Pagar.Cod_Rev IN ('DEVCOM', 'REBAJA') )")
                loComandoSeleccionar.AppendLine("				)										")
                loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_tip IN ('FACT', 'N/CR', 'N/DB')")
                loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Tip      BETWEEN " & lcParametro9Desde)
                loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
                loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status = 'Anulado' 										")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")

            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''Obtencion de las Ordenes de pago '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lcParametro6Desde.ToUpper() = "SI" Then

                loComandoSeleccionar.AppendLine(" UNION ALL")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine(" SELECT	")
                loComandoSeleccionar.AppendLine("			'Orden_Pago'													AS Tabla,		")
                loComandoSeleccionar.AppendLine("			'Orden de Pago' 												AS cod_tip,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Codigo_Tipo,	")
                loComandoSeleccionar.AppendLine("			CAST(Ordenes_Pagos.Documento AS CHAR(30))	 					AS Documento,	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Control	 										AS Control,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Referencia,	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Factura	 										AS Factura,		")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Status	 										AS Status,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Doc_Ori, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Documento_Afectado, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Exp_Importacion, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro	 										AS Cod_Pro, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini	 										AS Fec_Ini, 	")
                loComandoSeleccionar.AppendLine("			'Debito'														As Tip_Doc, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Bru	 										AS Mon_Bru, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net	 										AS Mon_Net, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Dis_Imp	 										AS Dis_Imp, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Des, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Rec, 	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS Por_Des, 	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS Por_Rec, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Otr1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Otr2, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    														AS Mon_Otr3, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp1, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp2, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp3, 	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe,	    ")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Com_Ret, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Ret											AS Mon_Ret,		")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_exe,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_bas,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_imp,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS Exonerado,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS No_Sujeto,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS Sin_Derecho_CF,	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Motivo                                            AS Comentario, ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro													")
                loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Nom_Pro = '')								")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Nom_Pro												")
                loComandoSeleccionar.AppendLine("					ELSE Ordenes_Pagos.Nom_Pro												")
                loComandoSeleccionar.AppendLine("				END)																		")
                loComandoSeleccionar.AppendLine("			END)															AS Nom_Pro, 	")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif														")
                loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Rif = '')									")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Rif													")
                loComandoSeleccionar.AppendLine("					ELSE Ordenes_Pagos.Rif													")
                loComandoSeleccionar.AppendLine("				END)																		")
                loComandoSeleccionar.AppendLine("			END)															AS Rif,			")
                loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pro												AS Tip_Pro,		")

                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '1'									        ")
                loComandoSeleccionar.AppendLine("				THEN 'NR' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '2'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'NNR' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '3'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'JD' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '4'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'JND' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '5'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'EXE' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE  																        ")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Cod_Per = '6'										    ")
                loComandoSeleccionar.AppendLine("				THEN 'TN' 																    ")
                loComandoSeleccionar.AppendLine("				ELSE 'OTR' 																    ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															                ")
                loComandoSeleccionar.AppendLine("			END)															    AS Cod_Per,	")

                loComandoSeleccionar.AppendLine("			(CASE WHEN Ordenes_Pagos.Status = 'Anulado'										")
                loComandoSeleccionar.AppendLine("				THEN '03-ANU' 																")
                loComandoSeleccionar.AppendLine("				ELSE '01-REG' 																")
                loComandoSeleccionar.AppendLine("			END)															AS Transaccion	")

                loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos ")
                loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro")
                loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Ipos = 0")
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Status = 'Confirmado'")

                If lcParametro7Desde.ToUpper = "NO" Then
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Mon_Imp <> 0")
                End If

                If lcParametro5Desde = "Igual" Then
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If

                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
                loComandoSeleccionar.AppendLine("")

            End If

            '**********************************************************************************************************
            '****************************** Expedientes de Importaciones, si aplican **********************************
            '**********************************************************************************************************
            loComandoSeleccionar.AppendLine("UPDATE	#tmpLibroCompras ")
            loComandoSeleccionar.AppendLine("SET    Exp_Importacion = Importaciones.Caracter8 ")
            loComandoSeleccionar.AppendLine("FROM	(SELECT	Cod_Reg, ")
            loComandoSeleccionar.AppendLine("               Caracter8 ")
            loComandoSeleccionar.AppendLine("       FROM    Campos_Extras ")
            loComandoSeleccionar.AppendLine("       WHERE	Origen = 'Compras') AS Importaciones ")
            loComandoSeleccionar.AppendLine("WHERE	Importaciones.Cod_Reg			=	#tmpLibroCompras.Documento ")
            loComandoSeleccionar.AppendLine("AND	#tmpLibroCompras.Codigo_Tipo	=	'FACT' ")


            '**********************************************************************************************************
            '****************************** Retenciones de IVA, si aplican ********************************************
            '**********************************************************************************************************
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UPDATE		#tmpLibroCompras ")
            loComandoSeleccionar.AppendLine("SET		Com_Ret = Retenciones.Com_Ret, ")
            loComandoSeleccionar.AppendLine("			Fec_Ret = Retenciones.Fec_Ret, ")
            loComandoSeleccionar.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT		#tmpLibroCompras.Documento		AS Documento,")
            loComandoSeleccionar.AppendLine("						#tmpLibroCompras.Cod_Tip		AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						retenciones_documentos.num_com	AS Com_Ret,")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.fec_ini			AS Fec_Ret,")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.mon_net			AS Mon_Ret")
            loComandoSeleccionar.AppendLine("			FROM		retenciones_documentos ")
            loComandoSeleccionar.AppendLine("				JOIN	#tmpLibroCompras")
            loComandoSeleccionar.AppendLine("					ON	#tmpLibroCompras.Documento = retenciones_documentos.doc_ori")
            loComandoSeleccionar.AppendLine("					AND	#tmpLibroCompras.Codigo_Tipo = retenciones_documentos.cla_ori")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loComandoSeleccionar.AppendLine("				JOIN	Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					ON	Cuentas_Pagar.documento = retenciones_documentos.doc_des")
            loComandoSeleccionar.AppendLine("					AND	Cuentas_Pagar.cod_tip = retenciones_documentos.cla_des")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("		) AS Retenciones")
            loComandoSeleccionar.AppendLine("WHERE	Retenciones.Documento = #tmpLibroCompras.documento")
            loComandoSeleccionar.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroCompras.Cod_Tip")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" ")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Exclusión de impuestos segun gaceta 6152")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpLibroCompras ")
            loComandoSeleccionar.AppendLine("SET     Impuesto_Excluir = Excluir.Mon_Imp1")
            loComandoSeleccionar.AppendLine("FROM (  SELECT X.Documento, X.Cod_Tip, SUM(X.Mon_Imp1) Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("        FROM (")
            loComandoSeleccionar.AppendLine("            SELECT      Cuentas_Pagar.Documento, Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("                        (CASE WHEN Cuentas_Pagar.Cod_Tip = 'N/CR' ")
            loComandoSeleccionar.AppendLine("                         THEN -Cuentas_Pagar.Mon_Imp1 ELSE Cuentas_Pagar.Mon_Imp1 END) Mon_Imp1")
            loComandoSeleccionar.AppendLine("            FROM        #tmpLibroCompras")
            loComandoSeleccionar.AppendLine("                JOIN    Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("                    ON  Cuentas_Pagar.Documento = #tmpLibroCompras.Documento")
            loComandoSeleccionar.AppendLine("                    AND Cuentas_Pagar.Cod_Tip = #tmpLibroCompras.Codigo_Tipo")
            loComandoSeleccionar.AppendLine("                    AND Cuentas_Pagar.Automatico = 0")
            loComandoSeleccionar.AppendLine("                    AND CAST(Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
            loComandoSeleccionar.AppendLine("            UNION ALL  ")
            loComandoSeleccionar.AppendLine("            SELECT      Compras.Documento, 'FACT' AS Cod_Tip, Compras.Mon_Imp1 Mon_Imp1")
            loComandoSeleccionar.AppendLine("            FROM        #tmpLibroCompras")
            loComandoSeleccionar.AppendLine("                JOIN    Compras ")
            loComandoSeleccionar.AppendLine("                    ON  Compras.Documento = #tmpLibroCompras.Documento")
            loComandoSeleccionar.AppendLine("            WHERE       CAST(Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
            loComandoSeleccionar.AppendLine("            UNION ALL  ")
            loComandoSeleccionar.AppendLine("            SELECT      Renglones_Compras.Documento, 'FACT' AS Cod_Tip, Renglones_Compras.Mon_Imp1 Mon_Imp1")
            loComandoSeleccionar.AppendLine("            FROM        #tmpLibroCompras")
            loComandoSeleccionar.AppendLine("                JOIN    Compras ")
            loComandoSeleccionar.AppendLine("                    ON  Compras.Documento = #tmpLibroCompras.Documento")
            loComandoSeleccionar.AppendLine("                    AND CAST(Compras.Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') <> 'true'")
            loComandoSeleccionar.AppendLine("                JOIN    Renglones_Compras ")
            loComandoSeleccionar.AppendLine("                    ON  Renglones_Compras.Documento = Compras.Documento")
            loComandoSeleccionar.AppendLine("                    AND CAST(Renglones_Compras.Not_Sta AS XML).value('(clasificacion/logico1)[1]', 'VARCHAR(MAX)') = 'true'")
            loComandoSeleccionar.AppendLine("        ) X ")
            loComandoSeleccionar.AppendLine("        GROUP BY X.Documento, X.Cod_Tip")
            loComandoSeleccionar.AppendLine("    ) Excluir")
            loComandoSeleccionar.AppendLine("WHERE   #tmpLibroCompras.Documento = Excluir.Documento")
            loComandoSeleccionar.AppendLine("    AND #tmpLibroCompras.Codigo_Tipo = Excluir.Cod_Tip;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT     ROW_NUMBER() OVER(ORDER BY Fec_Ini) AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           #tmpLibroCompras.*, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN (SUBSTRING(Comentario,1,12) = 'IMPORTACION') THEN SUBSTRING(Comentario,13,8) ELSE SPACE(8)	END Expediente ")
            loComandoSeleccionar.AppendLine("FROM       #tmpLibroCompras	")
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            Dim loDistribucion As System.Xml.XmlDocument
            Dim laImpuestos As System.Xml.XmlNodeList
            Dim loTabla As New DataTable()

            loTabla = laDatosReporte.Tables(0)

            For Each loFila As DataRow In loTabla.Rows

                If Not String.IsNullOrEmpty(Trim(loFila.Item("dis_imp"))) Then

                    loDistribucion = New System.Xml.XmlDocument()
                    Try

                        loDistribucion.LoadXml(Trim(loFila.Item("dis_imp")))

                    Catch ex As Exception

                        Continue For

                    End Try

                    laImpuestos = loDistribucion.SelectNodes("impuestos/impuesto")

                    If (loFila.Item("Cod_Tip").Equals("Orden de Pago")) Then

                        If laImpuestos.Count >= 1 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 2 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 3 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                    Else

                        'GACETA 5162: Si se debe excluir el monto total del impuesto: se asegura de que el campo
                        'Impuesto_Excluir tenga exactamente el monto de la distribución (para evitar error por redondeo)
                        Dim llExcluirTodoImpuesto As Boolean = 0
                        If CDec(loFila.item("Impuesto_Excluir")) = CDec(loFila.item("Mon_imp1")) Then
                            llExcluirTodoImpuesto = 1
                        End If

                        If laImpuestos.Count >= 1 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 2 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 3 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        'GACETA 5162: Documentos que se excluyen por completo
                        If llExcluirTodoImpuesto Then
                            loFila.Item("Impuesto_Excluir") = CDec(loFila.Item("mon_imp1")) + CDec(loFila.Item("mon_imp2")) + CDec(loFila.Item("mon_imp3"))
                        End If

                    End If

                    loFila.Item("subt_imp") = CDec(loFila.Item("mon_imp1")) + CDec(loFila.Item("mon_imp2")) + CDec(loFila.Item("mon_imp3"))
                    loFila.Item("subt_exe") = CDec(loFila.Item("mon_exe1")) + CDec(loFila.Item("mon_exe2")) + CDec(loFila.Item("mon_exe3"))
                    loFila.Item("subt_bas") = CDec(loFila.Item("mon_bas1")) + CDec(loFila.Item("mon_bas2")) + CDec(loFila.Item("mon_bas3"))

                End If

            Next loFila

            loTabla.AcceptChanges()

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (loTabla.Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            Dim laDatosReporte2 As New DataSet()
            Dim loTabla2 As New DataTable()

            loTabla2 = loTabla.Copy()
            loTabla.Dispose()
            loTabla = Nothing
            laDatosReporte.Dispose()
            laDatosReporte = Nothing
            laDatosReporte2.Tables.Add(loTabla2)
            laDatosReporte2.AcceptChanges()

            '-------------------------------------------------------------------------------------------'
            ' Carga la imagen del logo en cusReportes
            '-------------------------------------------------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte2.Tables(0), "LogoEmpresa")
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras_MHL", laDatosReporte2)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLibro_Compras_MHL.ReportSource = loObjetoReporte

            '-------------------------------------------------------------------------------------------'
            'Selección de opcion por excel (Microsoft Excel - xls)
            '-------------------------------------------------------------------------------------------'
            If (Me.Request.QueryString("salida").ToLower = "xls") Then

                '-------------------------------------------------------------------------------------------'
                ' Ruta donde se creara temporalmente el archivo
                '-------------------------------------------------------------------------------------------'
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rLibro_Compras_MHL" & Guid.NewGuid().ToString("N") & ".xls")
                '-------------------------------------------------------------------------------------------'
                ' Se exporta para crear el archivo temporal
                '-------------------------------------------------------------------------------------------'
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)
                '-------------------------------------------------------------------------------------------'
                ' Se modifica el contenido del archivo
                '-------------------------------------------------------------------------------------------'
                Me.modificar_excel(lcFileName, laDatosReporte)
                '-------------------------------------------------------------------------------------------'
                ' Se coloca en la respuesta para descargar
                '-------------------------------------------------------------------------------------------'
                Me.Response.Clear()
                Me.Response.Buffer = True
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rLibro_Compras_MHL.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()
                Me.Response.End()

            End If

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try


    End Sub

    Private Sub modificar_excel(ByVal loFileName As String, ByVal loDatosReporte As DataSet)

        Try

            ''-------------------------------------------------------------------------------------------'
            '' Se inicializa el objeto a la aplicacion excel
            ''-------------------------------------------------------------------------------------------'
            'loAppExcel = New Excel.Application()
            'loAppExcel.Visible = False
            'loAppExcel.DisplayAlerts = False

            ''-------------------------------------------------------------------------------------------'
            '' Se carga el archivo excel a modificar
            ''-------------------------------------------------------------------------------------------'
            'Dim lcLibrosExcel As Excel.Workbooks = loAppExcel.Workbooks
            'Dim lcLibroExcel As Excel.Workbook = lcLibrosExcel.Open(loFileName)

            ''-------------------------------------------------------------------------------------------'
            '' Se activa la primera hoja del libro donde se almacenara toda la informacion
            ''-------------------------------------------------------------------------------------------'
            'Dim lcHojaExcel As Excel.Worksheet
            'lcHojaExcel = lcLibroExcel.Worksheets(1)
            'lcHojaExcel.Activate()

            ''-------------------------------------------------------------------------------------------'
            '' Se selecciona toda la hoja para blanquera todo
            ''   - El número total de columnas disponibles en Excel
            ''       Viejo límite: 256           (2^8)       (Excel 2003 o inferor)
            ''       Nuevo límite: 16.384        (2^14)      (Excel 2007)
            ''   - El número total de filas disponibles en Excel
            ''       Viejo límite: 65.536        (2^16)      (Excel 2003 o inferior)
            ''       Nuevo límite: el 1.048.576  (2^20)      (Excel 2007)
            ''-------------------------------------------------------------------------------------------'
            'Dim lcRango As Excel.Range = lcHojaExcel.Range("A1:IV65536")
            'lcRango.Select()
            'lcRango.Clear()
            'lcRango.Font.Size = 9
            'lcRango.Font.Name = "Tahoma"

            ''-------------------------------------------------------------------------------------------'
            '' Nombre de la empresa
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(1, 1).Value = cusAplicacion.goEmpresa.pcNombre.ToUpper

            ''-------------------------------------------------------------------------------------------'
            '' Nombre del modulo
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(2, 1).Value = "Compras"

            ''-------------------------------------------------------------------------------------------'
            '' Titulo del reporte
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("A3:S3")
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.Value = Me.pcNombreReporte
            'lcRango.Font.Size = 14
            'lcRango.Font.Bold = True
            'lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            'lcRango.Rows.AutoFit()

            ''-------------------------------------------------------------------------------------------'
            '' Parametros del reporte
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("A4:S4")
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.Value = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
            'lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
            'lcRango.Rows.AutoFit()

            'lcRango = lcHojaExcel.Range("K5:S5")
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.Font.Bold = True
            'lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            ''-------------------------------------------------------------------------------------------'
            '' Membrete de la tabla de informacion del reporte
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(6, 2).Value = "Fecha"
            'lcHojaExcel.Cells(6, 3).Value = "RIF"
            'lcHojaExcel.Cells(6, 4).Value = "Nombre o"
            'lcHojaExcel.Cells(6, 5).Value = "Tipo de"
            'lcHojaExcel.Cells(6, 6).Value = "Comprobante"
            'lcHojaExcel.Cells(6, 7).Value = "Expediente"
            'lcHojaExcel.Cells(6, 8).Value = "Factura"
            'lcHojaExcel.Cells(6, 9).Value = "N° de"

            'lcHojaExcel.Cells(6, 10).Value = "Nº de"
            'lcHojaExcel.Cells(6, 11).Value = "Nota de"
            'lcHojaExcel.Cells(6, 12).Value = "Nota de"
            'lcHojaExcel.Cells(6, 13).Value = "Tipo de"
            'lcHojaExcel.Cells(6, 14).Value = "Total de"
            'lcHojaExcel.Cells(6, 15).Value = "Sin Derecho"
            'lcHojaExcel.Cells(6, 16).Value = "Base"
            'lcHojaExcel.Cells(6, 17).Value = "Porcentaje"
            'lcHojaExcel.Cells(6, 18).Value = "Monto de"
            'lcHojaExcel.Cells(6, 19).Value = "Impuesto"

            ''-------------------------------------------------------------------------------------------'
            '' Se le da formato a las celdas del membrete de la tabla
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("A6:S6")
            'lcRango.Select()
            'lcRango.Font.Bold = True

            ''-------------------------------------------------------------------------------------------'
            '' Membrete de la tabla de informacion del reporte
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(7, 4).Value = "Razón Social"
            'lcHojaExcel.Cells(7, 5).Value = "Persona"
            'lcHojaExcel.Cells(7, 6).Value = "Impuesto"
            'lcHojaExcel.Cells(7, 7).Value = "Importación"
            'lcHojaExcel.Cells(7, 9).Value = "Control"

            'lcHojaExcel.Cells(7, 10).Value = "Expediente"
            'lcHojaExcel.Cells(7, 11).Value = "Débito"
            'lcHojaExcel.Cells(7, 12).Value = "Crédito"
            'lcHojaExcel.Cells(7, 13).Value = "Trans."
            'lcHojaExcel.Cells(7, 14).Value = "Compras"
            'lcHojaExcel.Cells(7, 15).Value = "Crédito Fiscal"
            'lcHojaExcel.Cells(7, 16).Value = "Imponible"
            'lcHojaExcel.Cells(7, 17).Value = "Impuesto"
            'lcHojaExcel.Cells(7, 18).Value = "Impuesto"
            'lcHojaExcel.Cells(7, 19).Value = "Retenido"

            ''-------------------------------------------------------------------------------------------'
            '' Se le da formato a las celdas del membrete de la tabla
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("A7:S7")
            'lcRango.Select()
            'lcRango.Font.Bold = True

            ''-------------------------------------------------------------------------------------------'
            '' Formato a las columnas de la tabla de informacion del reporte
            ''-------------------------------------------------------------------------------------------'
            ''lcRango = lcHojaExcel.Range("B1")
            ''lcRango.EntireColumn.NumberFormat = "DD/MM/YYYY"
            ''lcRango = lcHojaExcel.Range("C1:J1")
            ''lcRango.EntireColumn.NumberFormat = "@"
            ''lcRango = lcHojaExcel.Range("A1:S1")
            ''lcRango.EntireColumn.NumberFormat = "###,###,###,##0.00"

            ''-------------------------------------------------------------------------------------------'
            '' Formato a las celdas de la fecha y hora de creacion
            '' Fecha y hora de creacion
            ''-------------------------------------------------------------------------------------------'
            ''Dim lcFechaCreacion As DateTime = DateTime.Now()
            ''lcHojaExcel.Cells(1, 30).NumberFormat = "@"
            ''lcHojaExcel.Cells(1, 30).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ''lcHojaExcel.Cells(1, 30).Value = lcFechaCreacion.ToString("dd/MM/yyyy")
            ''lcHojaExcel.Cells(2, 30).NumberFormat = "@"
            ''lcHojaExcel.Cells(2, 30).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ''lcHojaExcel.Cells(2, 30).Value = lcFechaCreacion.ToString("hh:mm:ss tt")

            ''-------------------------------------------------------------------------------------------'
            '' Recorrido de los datos de la consulta a la base de datos
            ''-------------------------------------------------------------------------------------------'
            ''-------------------------------------------------------------------------------------------'
            '' Numero de la fila de los datos obtenidos de la consulta a la base de datos
            ''-------------------------------------------------------------------------------------------'
            'Dim lcFila As Integer

            ''-------------------------------------------------------------------------------------------'
            '' Total de filas de los datos obtenidos de la consulta a la base de datos
            ''-------------------------------------------------------------------------------------------'
            'Dim lcTotalFilas As Integer = loDatosReporte.Tables(0).Rows.Count - 1

            ''-------------------------------------------------------------------------------------------'
            '' Datos de la fila de la consulta a la base de datos
            ''-------------------------------------------------------------------------------------------'
            'Dim lcDatosFila As DataRow

            ''-------------------------------------------------------------------------------------------'
            '' Control de agrupamiento 1 - (Tipo de Documento)(Tipo_Documento)
            ''-------------------------------------------------------------------------------------------'
            'Dim grupo1 As String = ""

            ''-------------------------------------------------------------------------------------------'
            '' Control de agrupamiento 2 - (Codigo del Tipo de Documento)(Codigo_Tipo)
            ''-------------------------------------------------------------------------------------------'
            'Dim grupo2 As String = ""

            ''-------------------------------------------------------------------------------------------'
            '' Numero de la fila en el documento excel
            ''-------------------------------------------------------------------------------------------'
            'Dim lcNumFila As Integer = 7

            ''-------------------------------------------------------------------------------------------'
            '' Numero de la fila en el documento excel, inicial para la sumatoria del total para el grupo
            '' de Control(2)
            ''-------------------------------------------------------------------------------------------'
            'Dim lcFilaIni As Integer = 0

            ''-------------------------------------------------------------------------------------------'
            '' Numero de la fila en el documento excel, final para la sumatoria del total para el grupo de
            '' Control(2)
            ''-------------------------------------------------------------------------------------------'
            'Dim lcFilaFin As Integer = 0

            ''-------------------------------------------------------------------------------------------'
            '' Construccion de la formula de total para el grupo de control 1
            ''-------------------------------------------------------------------------------------------'
            'Dim lcTotalDocumento As String = "= 0"

            ''-------------------------------------------------------------------------------------------'
            '' Numeros de Filas 
            ''-------------------------------------------------------------------------------------------'
            'lcNumFila = lcNumFila + 1

            ''-------------------------------------------------------------------------------------------'
            '' Recorriendo las filas de los datos de la consulta a la base de datos
            ''-------------------------------------------------------------------------------------------'
            'For lcFila = 0 To lcTotalFilas - 1

            '    '-------------------------------------------------------------------------------------------'
            '    ' Se extrae los datos de la fila
            '    '-------------------------------------------------------------------------------------------'
            '    lcDatosFila = loDatosReporte.Tables(0).Rows(lcFila)

            '    '-------------------------------------------------------------------------------------------'
            '    ' Si el Tipo_Documento ha cambiado
            '    '-------------------------------------------------------------------------------------------'
            '    If (Trim(lcDatosFila("Codigo_Tipo")) <> grupo1) Then

            '        '-------------------------------------------------------------------------------------------'
            '        ' Si no es la primera vez que ocurre el cambio
            '        '-------------------------------------------------------------------------------------------'
            '        If (grupo1 <> "") Then

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se coloca la etiqueta de total 
            '            '-------------------------------------------------------------------------------------------'
            '            lcRango = lcHojaExcel.Range("H" & CStr(lcNumFila) & ":M" & CStr(lcNumFila))
            '            lcRango.Select()
            '            lcRango.MergeCells = True
            '            lcRango.EntireRow.Font.Bold = True
            '            lcRango.Value = "Total del Libro"

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se coloca la formula de total de todas las columnas
            '            '-------------------------------------------------------------------------------------------'
            '            lcHojaExcel.Cells(lcNumFila, 11).Formula = lcTotalDocumento.Replace("[LETRA]", "N")
            '            lcHojaExcel.Cells(lcNumFila, 12).Formula = lcTotalDocumento.Replace("[LETRA]", "O")
            '            lcHojaExcel.Cells(lcNumFila, 13).Formula = lcTotalDocumento.Replace("[LETRA]", "P")
            '            lcHojaExcel.Cells(lcNumFila, 14).Formula = lcTotalDocumento.Replace("[LETRA]", "R")
            '            lcHojaExcel.Cells(lcNumFila, 15).Formula = lcTotalDocumento.Replace("[LETRA]", "S")

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se inicicializa la variable de control
            '            '-------------------------------------------------------------------------------------------'
            '            lcTotalDocumento = "= 0"
            '            lcNumFila = lcNumFila + 2

            '        End If

            '        '-------------------------------------------------------------------------------------------'
            '        ' Se almacena el dato  en la variable de control para el grupo 1
            '        '-------------------------------------------------------------------------------------------'
            '        grupo1 = Trim(lcDatosFila("Codigo_Tipo"))

            '        '-------------------------------------------------------------------------------------------'
            '        ' Se coloca el titulo del grupo
            '        '-------------------------------------------------------------------------------------------'
            '        lcRango = lcHojaExcel.Range("B" & CStr(lcNumFila) & ":D" & CStr(lcNumFila))
            '        lcRango.Select()
            '        lcRango.MergeCells = True
            '        lcRango.EntireRow.Font.Bold = True
            '        lcRango.EntireRow.RowHeight = 30
            '        lcRango.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            '        lcRango.EntireRow.NumberFormat = "@"
            '        lcHojaExcel.Cells(lcNumFila, 1).Value = grupo1

            '        If (grupo1 = "01") Then
            '            lcRango.Value = "Retenciones de I.V.A. a Documentos fuera del Período"
            '        Else
            '            lcRango.Value = "Documentos del Libro de Ventas"
            '        End If

            '        lcNumFila = lcNumFila + 1

            '    End If

            '        '-------------------------------------------------------------------------------------------'
            '        ' Si el Codigo_Tipo ha cambiado
            '        '-------------------------------------------------------------------------------------------'
            '        If (Trim(lcDatosFila("Codigo_Tipo")) <> grupo2) Then

            '        '-------------------------------------------------------------------------------------------'
            '        ' Si no es la primera vez que ocurre el cambio
            '        '-------------------------------------------------------------------------------------------'
            '        If (grupo2 <> "") Then

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se almacena el numero de la fila final del grupo de datos
            '            '-------------------------------------------------------------------------------------------'
            '            lcFilaFin = lcNumFila - 1

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se coloca la etiqueta de total 
            '            '-------------------------------------------------------------------------------------------'
            '            lcRango = lcHojaExcel.Range("H" & CStr(lcNumFila) & ":M" & CStr(lcNumFila))
            '            lcRango.Select()
            '            lcRango.MergeCells = True
            '            lcRango.EntireRow.Font.Bold = True
            '            lcRango.Value = "Total del Tipo de Documento"

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se coloca la formula de total de todas las columnas
            '            '-------------------------------------------------------------------------------------------'
            '            lcHojaExcel.Cells(lcNumFila, 11).Formula = "=SUM(N" & CStr(lcFilaIni) & ":N" & CStr(lcFilaFin) & ")"
            '            lcHojaExcel.Cells(lcNumFila, 12).Formula = "=SUM(O" & CStr(lcFilaIni) & ":O" & CStr(lcFilaFin) & ")"
            '            lcHojaExcel.Cells(lcNumFila, 13).Formula = "=SUM(P" & CStr(lcFilaIni) & ":P" & CStr(lcFilaFin) & ")"
            '            lcHojaExcel.Cells(lcNumFila, 14).Formula = "=SUM(R" & CStr(lcFilaIni) & ":R" & CStr(lcFilaFin) & ")"
            '            lcHojaExcel.Cells(lcNumFila, 15).Formula = "=SUM(S" & CStr(lcFilaIni) & ":S" & CStr(lcFilaFin) & ")"

            '            '-------------------------------------------------------------------------------------------'
            '            ' Se agrega la celda del total del grupo 2 a la formula del grupo 1
            '            '-------------------------------------------------------------------------------------------'
            '            lcTotalDocumento = lcTotalDocumento & " + [LETRA]" & CStr(lcNumFila)

            '            lcNumFila = lcNumFila + 2

            '        End If

            '        '-------------------------------------------------------------------------------------------'
            '        ' Se almacena el dato  en la variable de control para el grupo 2
            '        '-------------------------------------------------------------------------------------------'
            '        grupo2 = Trim(lcDatosFila("Codigo_Tipo"))

            '        '-------------------------------------------------------------------------------------------'
            '        ' Se coloc el titulo del grupo
            '        '-------------------------------------------------------------------------------------------'
            '        lcRango = lcHojaExcel.Range("B" & CStr(lcNumFila) & ":D" & CStr(lcNumFila))
            '        lcRango.Select()
            '        lcRango.MergeCells = True
            '        lcRango.EntireRow.Font.Bold = True
            '        lcRango.EntireRow.RowHeight = 20
            '        lcRango.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            '        lcRango.EntireRow.NumberFormat = "@"
            '        lcRango.Value = "Tipo de Documento: " & grupo2
            '        lcNumFila = lcNumFila + 1

            '        '-------------------------------------------------------------------------------------------'
            '        ' Se almacena el numero de la fila inicial del grupo de datos
            '        '-------------------------------------------------------------------------------------------'
            '        lcFilaIni = lcNumFila

            '    End If

            '    '-------------------------------------------------------------------------------------------'
            '    ' Se agrega la informacion en la tabla de los valores detallados
            '    '-------------------------------------------------------------------------------------------'
            '    lcHojaExcel.Cells(lcNumFila, 1).Value = lcDatosFila("Fec_Ini")
            '    lcHojaExcel.Cells(lcNumFila, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            '    lcHojaExcel.Cells(lcNumFila, 2).Value = lcDatosFila("Rif")
            '    lcHojaExcel.Cells(lcNumFila, 3).Value = lcDatosFila("Nom_Pro")
            '    lcHojaExcel.Cells(lcNumFila, 4).Value = lcDatosFila("Cod_per")
            '    lcHojaExcel.Cells(lcNumFila, 5).Value = lcDatosFila("Com_Ret")
            '    lcHojaExcel.Cells(lcNumFila, 6).Value = lcDatosFila("Exp_Importacion")
            '    lcHojaExcel.Cells(lcNumFila, 7).Value = lcDatosFila("Factura")
            '    lcHojaExcel.Cells(lcNumFila, 8).Value = lcDatosFila("Control")
            '    lcHojaExcel.Cells(lcNumFila, 9).Value = lcDatosFila("Expediente")

            '    ''-------------------------------------------------------------------------------------------'
            '    '' Bloque 1
            '    ''-------------------------------------------------------------------------------------------'
            '    'If (Trim(lcDatosFila("Codigo_Tipo")) = "N/DB") Then
            '    '    lcHojaExcel.Cells(lcNumFila, 10).Value = lcDatosFila("Documento")
            '    'Else
            '    '    lcHojaExcel.Cells(lcNumFila, 10).Value = New String("")
            '    'End If

            '    'If (Trim(lcDatosFila("Codigo_Tipo")) = "N/CR") Then
            '    '    lcHojaExcel.Cells(lcNumFila, 11).Value = lcDatosFila("Documento")
            '    'Else
            '    '    lcHojaExcel.Cells(lcNumFila, 11).Value = New String("")
            '    'End If

            '    'lcHojaExcel.Cells(lcNumFila, 12).Value = lcDatosFila("Transaccion")
            '    'lcHojaExcel.Cells(lcNumFila, 13).Value = lcDatosFila("Documento_Afectado")
            '    'lcHojaExcel.Cells(lcNumFila, 14).Value = lcDatosFila("Mon_Net")
            '    'lcHojaExcel.Cells(lcNumFila, 15).Value = lcDatosFila("Sin_Derecho_CF")
            '    'lcHojaExcel.Cells(lcNumFila, 16).Value = lcDatosFila("Mon_Bru")
            '    'lcHojaExcel.Cells(lcNumFila, 17).Value = lcDatosFila("Por_Imp1")
            '    'lcHojaExcel.Cells(lcNumFila, 18).Value = lcDatosFila("Mon_Imp1")
            '    'lcHojaExcel.Cells(lcNumFila, 19).Value = lcDatosFila("Mon_Ret")

            '    lcNumFila = lcNumFila + 1

            'Next lcFila

            ''-------------------------------------------------------------------------------------------'
            '' Se almacena el numero de la fila final del grupo de datos
            ''-------------------------------------------------------------------------------------------'
            'lcFilaFin = lcNumFila - 1

            ''-------------------------------------------------------------------------------------------'
            '' Se coloca la etiqueta de total 
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("H" & CStr(lcNumFila) & ":M" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Total del Tipo de Documento"

            ''-------------------------------------------------------------------------------------------'
            '' Se coloca la formula de total de todas las columnas
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(lcNumFila, 11).Formula = "=SUM(N" & CStr(lcFilaIni) & ":N" & CStr(lcFilaFin) & ")"
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = "=SUM(O" & CStr(lcFilaIni) & ":O" & CStr(lcFilaFin) & ")"
            'lcHojaExcel.Cells(lcNumFila, 13).Formula = "=SUM(P" & CStr(lcFilaIni) & ":P" & CStr(lcFilaFin) & ")"
            'lcHojaExcel.Cells(lcNumFila, 14).Formula = "=SUM(R" & CStr(lcFilaIni) & ":R" & CStr(lcFilaFin) & ")"
            'lcHojaExcel.Cells(lcNumFila, 15).Formula = "=SUM(S" & CStr(lcFilaIni) & ":S" & CStr(lcFilaFin) & ")"

            ''-------------------------------------------------------------------------------------------'
            '' Se agrega la celda del total del grupo 2 a la formula del grupo 1
            ''-------------------------------------------------------------------------------------------'
            'lcTotalDocumento = lcTotalDocumento & " + [LETRA]" & CStr(lcNumFila)
            'lcNumFila = lcNumFila + 2

            ''-------------------------------------------------------------------------------------------'
            '' Se coloca la etiqueta de total 
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("H" & CStr(lcNumFila) & ":M" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.MergeCells = True
            'lcRango.EntireRow.Font.Bold = True
            'lcRango.Value = "Total del Libro"

            ''-------------------------------------------------------------------------------------------'
            '' Se coloca la formula de total de todas las columnas
            ''-------------------------------------------------------------------------------------------'
            'lcHojaExcel.Cells(lcNumFila, 11).Formula = lcTotalDocumento.Replace("[LETRA]", "N")
            'lcHojaExcel.Cells(lcNumFila, 12).Formula = lcTotalDocumento.Replace("[LETRA]", "O")
            'lcHojaExcel.Cells(lcNumFila, 13).Formula = lcTotalDocumento.Replace("[LETRA]", "P")
            'lcHojaExcel.Cells(lcNumFila, 14).Formula = lcTotalDocumento.Replace("[LETRA]", "R")
            'lcHojaExcel.Cells(lcNumFila, 15).Formula = lcTotalDocumento.Replace("[LETRA]", "S")

            ''-------------------------------------------------------------------------------------------'
            '' Se agregan los totales del Documento
            ''-------------------------------------------------------------------------------------------'
            ' ''lcNumFila = lcNumFila + 3
            ' ''lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            ' ''lcRango.Select()
            ' ''lcRango.MergeCells = True
            ' ''lcRango.EntireRow.Font.Bold = True
            ' ''lcRango.Value = "Total Ventas"
            ' ''lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            ' ''lcNumFila = lcNumFila + 1
            ' ''lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            ' ''lcRango.Select()
            ' ''lcRango.MergeCells = True
            ' ''lcRango.EntireRow.Font.Bold = True
            ' ''lcRango.Value = "Ventas Gravadas"
            ' ''lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            ' ''lcNumFila = lcNumFila + 1
            ' ''lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            ' ''lcRango.Select()
            ' ''lcRango.MergeCells = True
            ' ''lcRango.EntireRow.Font.Bold = True
            ' ''lcRango.Value = "Total Exentas"
            ' ''lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            ' ''lcNumFila = lcNumFila + 1
            ' ''lcRango = lcHojaExcel.Range("I" & CStr(lcNumFila) & ":J" & CStr(lcNumFila))
            ' ''lcRango.Select()
            ' ''lcRango.MergeCells = True
            ' ''lcRango.EntireRow.Font.Bold = True
            ' ''lcRango.Value = "Impuestos"
            ' ''lcHojaExcel.Cells(lcNumFila, 11).Value = "0,00 %"
            ' ''lcHojaExcel.Cells(lcNumFila, 12).Formula = "= 0"

            ' ''-------------------------------------------------------------------------------------------'
            '' Ajustamos el tamaño de las columnas
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("B1:S" & CStr(lcNumFila))
            'lcRango.Select()
            'lcRango.Columns.AutoFit()

            ''-------------------------------------------------------------------------------------------'
            '' Seleccionamos la primera celda del libro
            ''-------------------------------------------------------------------------------------------'
            'lcRango = lcHojaExcel.Range("A1")
            'lcRango.Select()

            ''-------------------------------------------------------------------------------------------'
            '' Cerramos y liberamos recursos
            ''-------------------------------------------------------------------------------------------'
            'mLiberar(lcRango)
            'mLiberar(lcHojaExcel)

            ''-------------------------------------------------------------------------------------------'
            ''Guardamos los cambios del libro activo
            ''-------------------------------------------------------------------------------------------'
            'lcLibroExcel.Close(True, loFileName)
            'mLiberar(lcLibroExcel)
            'loAppExcel.Application.Quit()
            'mLiberar(loAppExcel)

        Catch loExcepcion As Exception

            Me.mEscribirConsulta(loExcepcion.Message)
            Me.Response.Flush()
            Me.Response.Close()
            Me.Response.End()

        Finally

            '-------------------------------------------------------------------------------------------'
            ' Se forza el cierre del proceso excel
            '-------------------------------------------------------------------------------------------'
            'Dim Lista_Procesos() As Diagnostics.Process
            'Dim p As Diagnostics.Process
            'Lista_Procesos = Diagnostics.Process.GetProcessesByName("EXCEL")
            'For Each p In Lista_Procesos
            '    Try
            '        p.Kill()
            '    Catch
            '    End Try
            'Next

            'GC.Collect()

        End Try

    End Sub

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 15/04/13: Codigo inicial, a partir de rLibro_Compras_GrupoAlegria					'
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Agregado filtro para incluir solo retenciones de IVA (no ISLR ni Patente). '
'-------------------------------------------------------------------------------------------'
' RJG: 03/06/13: Se cambió la columna "Subtotal Factura" (Monto Bruto) por "Compras sin     '
'                Impuesto".                                                                 '
'-------------------------------------------------------------------------------------------'
' RJG: 30/08/13: Se ajustó el cálculo de totales para descargar los montos de documentos    '
'                anulados.                                                                  '
'-------------------------------------------------------------------------------------------'
' JJD: 17/01/14: Ajustado para ejecutarlo en Michelin                                       '
'-------------------------------------------------------------------------------------------'
' JJD: 21/01/14: Continuacion del ajuste para Michelin                                      '
'-------------------------------------------------------------------------------------------'
' JJD: 22/01/14: Continuacion del ajuste para Michelin                                      '
'-------------------------------------------------------------------------------------------'
' RJG: 20/04/15: Se agregó la exclusión de los impuestos de compras/renglones que están     '
'                marcados como excluidos en la clasificación del documento (GACETA 6152).   '
'-------------------------------------------------------------------------------------------'
