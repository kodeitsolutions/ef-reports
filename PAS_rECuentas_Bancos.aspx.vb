'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "PAS_rECuentas_Bancos"
'-------------------------------------------------------------------------------------------'
Partial Class PAS_rECuentas_Bancos
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@CodCue_Desde	AS VARCHAR(15)")
            loComandoSeleccionar.AppendLine("DECLARE	@CodCue_Hasta	AS VARCHAR(15)")
            loComandoSeleccionar.AppendLine("DECLARE	@CodCon_Desde	AS VARCHAR(5)")
            loComandoSeleccionar.AppendLine("DECLARE	@CodCon_Hasta	AS VARCHAR(5)")
            loComandoSeleccionar.AppendLine("DECLARE	@TipMov_Desde	AS VARCHAR(15)")
            loComandoSeleccionar.AppendLine("DECLARE	@TipMov_Hasta	AS VARCHAR(15)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@FecIni          = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET	@FecFin          = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("SET	@CodCue_Desde    = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@CodCue_Hasta    = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET	@CodCon_Desde    = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET	@CodCon_Hasta    = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET	@TipMov_Desde    = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("SET	@TipMov_Hasta    = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero		AS DECIMAL(28, 10)	;")
            loComandoSeleccionar.AppendLine("SET @lnCero			= 0")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT     Movimientos_Cuentas.Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("			(SUM(Movimientos_Cuentas.Mon_Deb)- SUM(Movimientos_Cuentas.Mon_Hab)) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO		#tempSALDOINICIAL     ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue     ")
            loComandoSeleccionar.AppendLine("	JOIN	Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban     ")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Fec_Ini < @FecIni  ")
            loComandoSeleccionar.AppendLine("   AND		Movimientos_Cuentas.Cod_Cue BETWEEN @CodCue_Desde AND @CodCue_Hasta")
            loComandoSeleccionar.AppendLine("   AND		Movimientos_Cuentas.Cod_Con BETWEEN @CodCon_Desde AND @CodCon_Hasta")
            loComandoSeleccionar.AppendLine("   AND		Movimientos_Cuentas.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue ")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT		0										AS Orden, 		")
            loComandoSeleccionar.AppendLine("			Movimientos_Cuentas.Cod_Cue				AS Cod_Cue,		")
            loComandoSeleccionar.AppendLine("        	Cuentas_Bancarias.Num_Cue				AS Num_Cue,		")
            loComandoSeleccionar.AppendLine("        	Bancos.Nom_Ban							AS Nom_Ban,		")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Fec_Ini				AS Fec_Ini,		")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Documento			AS Documento,	")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Cod_Tip				AS Cod_Tip,     ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Tip_Doc				AS Tip_Doc,     ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Comentario			AS Comentario,  ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Tip_Ori				AS Tip_Ori,     ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Mon_Deb				AS Mon_Deb,     ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Mon_Hab				AS Mon_Hab,     ")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Mon_Imp1			AS Mon_Imp1,	")
            loComandoSeleccionar.AppendLine("        	Movimientos_Cuentas.Referencia			AS Referencia,	")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Sal_Ini,		")
            loComandoSeleccionar.AppendLine("			@lnCero									AS Mon_Sal		")
            loComandoSeleccionar.AppendLine("INTO		#tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Bancarias ON	Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Cod_Cue BETWEEN @CodCue_Desde AND @CodCue_Hasta")
            loComandoSeleccionar.AppendLine("	JOIN	Bancos ON	Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Detalles_Cobros ON	Movimientos_Cuentas.Doc_Ori = Detalles_Cobros.Documento")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Cobros'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_Cobros.Tip_Des = 'Movimientos_Cuentas'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_Cobros.Doc_Des = Movimientos_Cuentas.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Detalles_Pagos ON	Movimientos_Cuentas.Doc_Ori = Detalles_Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_Pagos.Tip_Des = 'Movimientos_Cuentas'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_Pagos.Doc_Des = Movimientos_Cuentas.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Detalles_oPagos ON	Movimientos_Cuentas.Doc_Ori = Detalles_oPagos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_oPagos.Tip_Des = 'Movimientos_Cuentas'")
            loComandoSeleccionar.AppendLine("		AND	Detalles_oPagos.Doc_Des = Movimientos_Cuentas.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Depositos ON	Movimientos_Cuentas.Doc_Ori = Depositos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Depositos'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Cobrar ON	Movimientos_Cuentas.Doc_Ori = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Cobrar.Cod_Tip = 'CHEQ'")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Cuentas_Cobrar'")
            loComandoSeleccionar.AppendLine("WHERE	Movimientos_Cuentas.Cod_Con BETWEEN @CodCon_Desde AND @CodCon_Hasta")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini BETWEEN @FecIni AND @FecFin")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Tip BETWEEN @TipMov_Desde AND @TipMov_Hasta")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tempMOVIMIENTO")
            loComandoSeleccionar.AppendLine("SET		Orden = M.Orden,")
            loComandoSeleccionar.AppendLine("		Mon_Sal = M.Mon_Deb-M.Mon_Hab,")
            loComandoSeleccionar.AppendLine("		Sal_Ini = M.Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	ROW_NUMBER() ")
            loComandoSeleccionar.AppendLine("						OVER (	PARTITION BY #tempMOVIMIENTO.Cod_Cue")
            loComandoSeleccionar.AppendLine("								ORDER BY #tempMOVIMIENTO.Fec_Ini, (CASE WHEN #tempMOVIMIENTO.Cod_Tip='' THEN 'zzzzzzzzz' ELSE #tempMOVIMIENTO.Cod_Tip END )  ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("					#tempMOVIMIENTO.Cod_Tip, #tempMOVIMIENTO.Documento,")
            loComandoSeleccionar.AppendLine("					ISNULL(SI.Sal_Ini, @lnCero) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("					#tempMOVIMIENTO.Mon_Deb AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("					#tempMOVIMIENTO.Mon_Hab AS Mon_Hab")
            loComandoSeleccionar.AppendLine("			FROM	#tempMOVIMIENTO			")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Cue, SUM(Sal_Ini) AS Sal_Ini FROM #tempSALDOINICIAL GROUP BY Cod_Cue) AS SI")
            loComandoSeleccionar.AppendLine("				ON SI.Cod_Cue = #tempMOVIMIENTO.Cod_Cue")
            loComandoSeleccionar.AppendLine("		) AS M		")
            loComandoSeleccionar.AppendLine("WHERE	M.Cod_Tip = #tempMOVIMIENTO.Cod_Tip")
            loComandoSeleccionar.AppendLine("	AND	M.Documento = #tempMOVIMIENTO.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Cod_Cue, A.Num_Cue, A.Nom_Ban, A.Fec_Ini, A.Documento, A.Cod_Tip,  ")
            loComandoSeleccionar.AppendLine("		A.Tip_Doc, A.Comentario, A.Tip_Ori, A.Mon_Deb, A.Mon_Hab, A.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("		A.Referencia, A.Sal_Ini, SUM(B.Mon_Sal) +  A.Sal_Ini AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("       @FecIni AS Desde, @FecFin AS Hasta")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tempMOVIMIENTO AS B")
            loComandoSeleccionar.AppendLine("		ON B.Cod_Cue = A.Cod_Cue")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Cod_Cue, A.Num_Cue, A.Nom_Ban, A.Cod_Tip, A.Tip_Doc,")
            loComandoSeleccionar.AppendLine("		A.Documento, A.Fec_Ini, A.Referencia, A.Sal_Ini, A.Tip_Ori,")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab, A.Comentario, A.Mon_Imp1")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue ASC, A.Fec_Ini ASC, (CASE WHEN A.Cod_Tip='' THEN 'zzzzzzzzz' ELSE A.Cod_Tip END ) ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempSALDOINICIAL")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempMOVIMIENTO")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("PAS_rECuentas_Bancos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvPAS_rECuentas_Bancos.ReportSource = loObjetoReporte

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
' CMS: 11/06/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' DLC: 20/05/10: Se agregó a la consulta los campos de la fecha inicial del detalle de pago,'
'				 asi como tambien la referencia a este documento							'
'-------------------------------------------------------------------------------------------'
' RJG: 02/09/10: Modificado para diferenciar los 5 posibles origenes de un movimiento		'
'				 bancario.																	'	
'-------------------------------------------------------------------------------------------'
' RJG: 05/01/12: Corrección en la unión: Aparecían movimientos duplicados cuando el cobro/	'
'				 Pago/Orden de Pago tenía más de una forma de pago.							'
'-------------------------------------------------------------------------------------------'
' RJG: 31/01/12: Corrección en la unión: faltó aplicar el filtro de Cuenta Bancaria en uno	'
'				 los JOINs.																	'
'-------------------------------------------------------------------------------------------'
