'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rLibro_Ventas"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rLibro_Ventas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            'Dim Empresa As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@sp_FecIni          = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_FecFin          = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() OVER (ORDER BY CAST(Registros.Fec_Ini AS DATE), Registros.Documento ASC) AS Num,*")
            loComandoSeleccionar.AppendLine("FROM(")
            loComandoSeleccionar.AppendLine("		/*Facturas*/")
            loComandoSeleccionar.AppendLine("		SELECT		Cuentas_Cobrar.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Control									AS Control,")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Fec_Ini									AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("					THEN 'ANULADO' ")
            loComandoSeleccionar.AppendLine("					ELSE (CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("                               THEN  Clientes.Nom_Cli")
            loComandoSeleccionar.AppendLine("                               ELSE Cuentas_Cobrar.Nom_Cli END)")
            loComandoSeleccionar.AppendLine("			END)													AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("            (CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("            		THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("            		ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("            END)													AS Rif,")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Cod_Tip									AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("            (CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("            		THEN '***ANULADA***' ELSE Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            END)													AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Des END						AS Mon_Des,")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Rec END						AS Mon_Rec,")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Por_Des  								AS Por_Des,")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Por_Rec  								AS Por_Rec,")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr1 END					AS Mon_Otr1,")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr2 END					AS Mon_Otr2,")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr3 END					AS Mon_Otr3,")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Net END						AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) END AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Exe END						AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Cod_Imp									AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Por_Imp1 								AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("            CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Imp1 END					AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("            ''                                                     AS Com_Ret, ")
            loComandoSeleccionar.AppendLine("            Cuentas_Cobrar.Fec_Ini                                 AS Fec_Ret,")
            loComandoSeleccionar.AppendLine("			0                                                       AS Base_Ret, ")
            loComandoSeleccionar.AppendLine("           0														AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("           MONTH(" & lcParametro0Desde & " )				        AS Mes,")
            loComandoSeleccionar.AppendLine("           YEAR(" & lcParametro0Hasta & " )				        AS Anio")
            loComandoSeleccionar.AppendLine("			FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("				JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("			WHERE		Cuentas_Cobrar.Cod_Tip      =   'FACT' ")
            loComandoSeleccionar.AppendLine("							AND Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("			/*Facturas anuladas*/")
            loComandoSeleccionar.AppendLine("			UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("			SELECT		Facturas.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("						Facturas.Control								AS Control,")
            loComandoSeleccionar.AppendLine("						Facturas.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("						'ANULADO'                           			AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("						(CASE	WHEN Facturas.Rif = ''")
            loComandoSeleccionar.AppendLine("            					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("            					ELSE Facturas.Rif")
            loComandoSeleccionar.AppendLine("						END)											AS Rif,")
            loComandoSeleccionar.AppendLine("						'FACT'      									AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("						'***ANULADA***'              					AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("						Facturas.mon_des1 * 0							AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("						Facturas.mon_rec1 * 0							AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("						Facturas.Por_Des1  								AS Por_Des, ")
            loComandoSeleccionar.AppendLine("						Facturas.Por_Rec1  								AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Otr1 * 0   						AS Mon_Otr1,")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Otr2 * 0   						AS Mon_Otr2,")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Otr3 * 0							AS Mon_Otr3,")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Net  * 0							AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Bru  * 0 							AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Exe  * 0							AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("						Facturas.cod_imp1								AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("						Facturas.Por_Imp1								AS Por_Imp1,")
            loComandoSeleccionar.AppendLine("						Facturas.Mon_Imp1 * 0							AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("						''                                              AS Com_Ret, ")
            loComandoSeleccionar.AppendLine("						Facturas.Fec_Ini                                AS Fec_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                               AS Base_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                               AS Mon_Ret, ")
            loComandoSeleccionar.AppendLine("                       MONTH(" & lcParametro0Desde & " )				AS Mes,")
            loComandoSeleccionar.AppendLine("                       YEAR(" & lcParametro0Hasta & " )				AS Anio")
            loComandoSeleccionar.AppendLine("			FROM		Facturas")
            loComandoSeleccionar.AppendLine("				JOIN	Clientes ON Facturas.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("			WHERE		Facturas.Status      =   'Anulado' ")
            loComandoSeleccionar.AppendLine("							AND Facturas.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("			/*Notas de Débito*/")
            loComandoSeleccionar.AppendLine("			UNION ALL")
            loComandoSeleccionar.AppendLine("            ")
            loComandoSeleccionar.AppendLine("			SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Control								AS Control,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("						(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("            					THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("            					ELSE Cuentas_Cobrar.Nom_Cli")
            loComandoSeleccionar.AppendLine("						END)												AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("						(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("            					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("            					ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("						END)												AS Rif,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("            					THEN '***ANULADA***' ELSE Cuentas_Cobrar.Referencia")
            loComandoSeleccionar.AppendLine("						END)											AS Doc_Ori,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Des  							AS Mon_Des,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Rec  							AS Mon_Rec,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr1 							AS Mon_Otr1,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr2 							AS Mon_Otr2,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr3 							AS Mon_Otr3,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Net 							    AS Mon_Net,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Bru                              AS Mon_Bru,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Exe  							AS Mon_Exe,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,  ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Imp1 							AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("						''                                                  AS Com_Ret, ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Fec_Ini                              AS Fec_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                                   AS Base_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                                   AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("                       MONTH(" & lcParametro0Desde & " )				    AS Mes,")
            loComandoSeleccionar.AppendLine("                       YEAR(" & lcParametro0Hasta & " )				    AS Anio")
            loComandoSeleccionar.AppendLine("			FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("				JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("			WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/DB'	")
            loComandoSeleccionar.AppendLine("							AND Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("			/*Notas de crédito*/")
            loComandoSeleccionar.AppendLine("			UNION ALL")
            loComandoSeleccionar.AppendLine("			SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Control								AS Control,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("						(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("            					THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("            					ELSE Cuentas_Cobrar.Nom_Cli")
            loComandoSeleccionar.AppendLine("						END)												AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("						(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("            					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("            					ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("						END)												AS Rif,")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("            					THEN '***ANULADA***' ELSE Cuentas_Cobrar.Referencia")
            loComandoSeleccionar.AppendLine("						END)												AS Doc_Ori,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Des*(-1)  						AS Mon_Des,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Rec*(-1)  						AS Mon_Rec,   ")
            loComandoSeleccionar.AppendLine("					    Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loComandoSeleccionar.AppendLine("					    Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr1*(-1) 						AS Mon_Otr1,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr2*(-1) 						AS Mon_Otr2,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Otr3*(-1) 						AS Mon_Otr3,    ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Net*(-1)                         AS Mon_Net,     ")
            loComandoSeleccionar.AppendLine("						(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)*(-1) AS Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Exe*(-1) 						AS Mon_Exe,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,  ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Mon_Imp1*(-1)						AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("						''                                                  AS Com_Ret, ")
            loComandoSeleccionar.AppendLine("						Cuentas_Cobrar.Fec_Ini                              AS Fec_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                                   AS Base_Ret, ")
            loComandoSeleccionar.AppendLine("						0                                                   AS Mon_Ret,      ")
            loComandoSeleccionar.AppendLine("                       MONTH(" & lcParametro0Desde & " )				    AS Mes,")
            loComandoSeleccionar.AppendLine("                       YEAR(" & lcParametro0Hasta & " )				    AS Anio")
            loComandoSeleccionar.AppendLine("			FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("				JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("			WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            loComandoSeleccionar.AppendLine("							AND Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin)Registros")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")



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

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rLibro_Ventas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rLibro_Ventas.ReportSource = loObjetoReporte

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
' RJG: 28/02/13: Codigo inicial, a partir de rLibro_Ventas.aspx.							'
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Agregado filtro para incluir solo retenciones de IVA (no ISLR ni Patente). '
'-------------------------------------------------------------------------------------------'
' RJG: 29/07/13: Se agregaron las Facturas de Venta Anuladas. También se mostrarán los      '
'                montos de los documetnos anulados, pero sin contarlos para los totales. Se '
'                ajustaron los porcentajes de impuesto en el total para que muestre todos.  '
'-------------------------------------------------------------------------------------------'
