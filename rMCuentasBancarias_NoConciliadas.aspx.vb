'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCuentasBancarias_NoConciliadas"
'-------------------------------------------------------------------------------------------'
Partial Class rMCuentasBancarias_NoConciliadas

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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine("SELECT           ")
            loComandoSeleccionar.AppendLine("			Movimientos_Cuentas.Cod_Cue AS Cod_Cue,     ")
            loComandoSeleccionar.AppendLine("			(SUM(Movimientos_Cuentas.Mon_Deb) ")
            loComandoSeleccionar.AppendLine("			- SUM(Movimientos_Cuentas.Mon_Hab)) AS Saldo_Inicial  ")
            loComandoSeleccionar.AppendLine("INTO	#tmpSaldo_Inicial")
            loComandoSeleccionar.AppendLine("FROM	Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("WHERE      Movimientos_Cuentas.Doc_Cil     =   '0' ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Fec_Ini < " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Cue     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Mon     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status      IN  (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Tip   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Rev     Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Suc     Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Movimientos_Cuentas.Cod_Cue   ")
            
            
            

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("			ROW_NUMBER() OVER (PARTITION BY  Movimientos_Cuentas.Cod_Cue ORDER BY Movimientos_Cuentas.Cod_Cue ASC) AS Indice,")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpSaldo_Inicial.Saldo_Inicial,0) AS Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Imp1 AS IDB, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Tipos_Movimientos.Nom_Tip,")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban ")
            loComandoSeleccionar.AppendLine(" INTO #tmpTemporal")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN #tmpSaldo_Inicial ON (Movimientos_Cuentas.Cod_Cue    =   #tmpSaldo_Inicial.Cod_Cue)")
            loComandoSeleccionar.AppendLine(" JOIN Tipos_Movimientos ON (Movimientos_Cuentas.Cod_Tip	=   Tipos_Movimientos.Cod_Tip) ")
            loComandoSeleccionar.AppendLine(" JOIN Cuentas_Bancarias ON (Movimientos_Cuentas.Cod_Cue    =   Cuentas_Bancarias.Cod_Cue)")
            loComandoSeleccionar.AppendLine(" JOIN Bancos ON (Bancos.Cod_Ban    =   Cuentas_Bancarias.Cod_Ban)")
            loComandoSeleccionar.AppendLine(" JOIN Conceptos ON (Movimientos_Cuentas.Cod_Con     =   Conceptos.Cod_Con)")
            loComandoSeleccionar.AppendLine(" WHERE      Movimientos_Cuentas.Doc_Cil     =   '0' ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Fec_Ini     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Cue     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Mon     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Status      IN  (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Tip   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Rev     Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Suc     Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
			
			
			loComandoSeleccionar.AppendLine(" SELECT	#tmpTemporal.Documento,")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Indice,  ")
			loComandoSeleccionar.AppendLine("			CASE WHEN Indice = 1 THEN #tmpTemporal.Saldo_Inicial ELSE 0 END AS Saldo_Inicial,")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Cod_Tip, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Status, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Referencia, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Comentario, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Cod_Con, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Cod_Mon,  ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Mon_Deb, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Mon_Hab, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Tip_Ori, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Doc_Ori, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.IDB, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Tipo, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Nom_Tip,")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Num_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Nom_Con, ")
			loComandoSeleccionar.AppendLine("			#tmpTemporal.Nom_Ban")
			loComandoSeleccionar.AppendLine(" FROM #tmpTemporal")
			loComandoSeleccionar.AppendLine("ORDER BY   Cod_Cue, " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMCuentasBancarias_NoConciliadas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCuentasBancarias_NoConciliadas.ReportSource = loObjetoReporte

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
' MAT: 20/07/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
