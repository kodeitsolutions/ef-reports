'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rECuentas_Bancos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rECuentas_Bancos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            'Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            'Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("DECLARE @lcCodCue_Desde	AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodCue_Hasta	AS VARCHAR(10) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde	AS DATE = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta	AS DATE = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Movimientos_Cuentas.Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("		(SUM(Movimientos_Cuentas.Mon_Deb)- SUM(Movimientos_Cuentas.Mon_Hab)) AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO #tempSALDOINICIAL     ")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue     ")
            loComandoSeleccionar.AppendLine("	JOIN Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban     ")
            loComandoSeleccionar.AppendLine("WHERE Movimientos_Cuentas.Fec_Ini < @ldFecha_Desde  ")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCodCue_Desde AND @lcCodCue_Hasta")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("GROUP BY Movimientos_Cuentas.Cod_Cue                 ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Movimientos_Cuentas.Cod_Cue,		")
            loComandoSeleccionar.AppendLine("        Cuentas_Bancarias.Num_Cue,			")
            loComandoSeleccionar.AppendLine("        Bancos.Nom_Ban,					")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Fec_Ini,		")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Documento,		")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Cod_Tip,     	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Tip_Doc,     	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Comentario,  	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Tip_Ori,     	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Deb,     	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Hab,     	")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Mon_Imp1,		")
            loComandoSeleccionar.AppendLine("        Movimientos_Cuentas.Referencia,		")
            loComandoSeleccionar.AppendLine("		CASE Movimientos_Cuentas.Tip_Ori				")
            loComandoSeleccionar.AppendLine("			WHEN 'Cobros'			THEN Detalles_Cobros.Fec_Ini	")
            loComandoSeleccionar.AppendLine("			WHEN 'Pagos'			THEN Detalles_Pagos.Fec_Ini		")
            loComandoSeleccionar.AppendLine("			WHEN 'Ordenes_Pagos'	THEN Detalles_oPagos.Fec_Ini	")
            loComandoSeleccionar.AppendLine("			WHEN 'Depositos'		THEN Depositos.Fec_Ini			")
            loComandoSeleccionar.AppendLine("			WHEN 'Cuentas_Cobrar'	THEN Cuentas_Cobrar.Fec_Ini		")
            loComandoSeleccionar.AppendLine("		ELSE NULL									")
            loComandoSeleccionar.AppendLine("		END							AS Fec_Ini_Detalles")
            loComandoSeleccionar.AppendLine("INTO #tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cuentas     ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Movimientos_Cuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Cod_Cue BETWEEN @lcCodCue_Desde AND @lcCodCue_Hasta")
            loComandoSeleccionar.AppendLine("	JOIN Bancos ON Bancos.Cod_Ban = Cuentas_Bancarias.Cod_Ban")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Detalles_Cobros ON Movimientos_Cuentas.Doc_Ori = Detalles_Cobros.Documento")
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
            loComandoSeleccionar.AppendLine("		AND	Movimientos_Cuentas.Tip_Ori = 'Cuentas_Cobrar'		")
            loComandoSeleccionar.AppendLine("WHERE Movimientos_Cuentas.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'            ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tempMOVIMIENTO.Cod_Cue,   	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Num_Cue,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Nom_Ban,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Ini,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Documento, 		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Cod_Tip,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Tip_Doc,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Comentario,		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Tip_Ori,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Deb,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Hab,   		")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Mon_Imp1,		")
            loComandoSeleccionar.AppendLine("		ISNULL(#tempSALDOINICIAL.Sal_Ini,0) AS Sal_Ini,     ")
            loComandoSeleccionar.AppendLine("		0 AS Sal_Doc,     ")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Referencia,     ")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO.Fec_Ini_Detalles     ")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO     ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	#tempSALDOINICIAL ON #tempSALDOINICIAL.Cod_Cue = #tempMOVIMIENTO.Cod_Cue")
            loComandoSeleccionar.AppendLine("ORDER BY #tempMOVIMIENTO.Cod_Cue, #tempMOVIMIENTO.Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempSALDOINICIAL")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempMOVIMIENTO")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmetne los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Cue", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Num_Cue", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Ban", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Documento", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Tip_Doc", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna.MaxLength = 500
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Tip_Ori", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Imp1", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Referencia", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini_Detalles", GetType(String))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Cuenta_Actual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loNuevaFila.Item("Cod_Cue") = loFila("Cod_Cue")
                loNuevaFila.Item("Num_Cue") = loFila("Num_Cue")
                loNuevaFila.Item("Nom_Ban") = loFila("Nom_Ban")
                loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                loNuevaFila.Item("Documento") = loFila("Documento")
                loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                loNuevaFila.Item("Tip_Doc") = loFila("Tip_Doc")
                loNuevaFila.Item("Comentario") = loFila("Comentario")
                loNuevaFila.Item("Tip_Ori") = loFila("Tip_Ori")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Mon_Imp1") = loFila("Mon_Imp1")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                loNuevaFila.Item("Referencia") = loFila("Referencia")

                Dim ldFechaHoy As Date = Date.Today()
                If (loFila("Fec_Ini_Detalles") Is System.DBNull.Value) Then
                    loNuevaFila.Item("Fec_Ini_Detalles") = loNuevaFila.Item("Fec_Ini")
                Else
                    loNuevaFila.Item("Fec_Ini_Detalles") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini_Detalles")), "dd/MM/yyyy")
                End If

                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Cuenta_Actual = loFila("Cod_Cue")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Cod_Cue") <> Cuenta_Actual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Cod_Cue") = loFila("Cod_Cue")
                    loNuevaFila.Item("Num_Cue") = loFila("Num_Cue")
                    loNuevaFila.Item("Nom_Ban") = loFila("Nom_Ban")
                    loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                    loNuevaFila.Item("Documento") = loFila("Documento")
                    loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                    loNuevaFila.Item("Tip_Doc") = loFila("Tip_Doc")
                    loNuevaFila.Item("Comentario") = loFila("Comentario")
                    loNuevaFila.Item("Tip_Ori") = loFila("Tip_Ori")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Mon_Imp1") = loFila("Mon_Imp1")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    loNuevaFila.Item("Referencia") = loFila("Referencia")

                    If (loFila("Fec_Ini_Detalles") Is System.DBNull.Value) Then
                        loNuevaFila.Item("Fec_Ini_Detalles") = loNuevaFila.Item("Fec_Ini")
                    Else
                        loNuevaFila.Item("Fec_Ini_Detalles") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini_Detalles")), "dd/MM/yyyy")
                    End If

                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Cuenta_Actual = loFila("Cod_Cue")

                    loTabla.AcceptChanges()

                Next lnNumeroFila

                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'

                Me.mCargarLogoEmpresa(loDatosReporteFinal.Tables(0), "LogoEmpresa")

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rECuentas_Bancos", loDatosReporteFinal)
            Else
                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rECuentas_Bancos", laDatosReporte)
            End If

            
            loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
            'loObjetoReporte.ReportDefinition.ReportObjects("text8").Height = 0
            'loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
            'loObjetoReporte.ReportDefinition.ReportObjects("text8").Top = 0
            'loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rECuentas_Bancos.ReportSource = loObjetoReporte

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
