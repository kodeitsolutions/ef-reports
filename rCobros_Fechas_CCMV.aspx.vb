'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Fechas_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Fechas_CCMV
    Inherits vis2formularios.frmReporte

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
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cobros.Documento, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Bru    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Rec    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Des    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Imp1   * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Cuentas_Cobrar.Mon_Net    * (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN -1 ELSE 1 END)))    AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpTemporalCobros ")
            loComandoSeleccionar.AppendLine(" FROM		Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" WHERE		Cobros.Documento                =   Renglones_Cobros.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip      =   Renglones_Cobros.Cod_Tip ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Documento    =   Renglones_Cobros.Doc_Ori ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Cli      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Mon      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Ven      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Status       IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Zon    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc      Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  Cobros.Documento ")


            loComandoSeleccionar.AppendLine(" SELECT    #tmpTemporalCobros.Documento, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Bru As Mon_Bru02, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Rec, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Des, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 			#tmpTemporalCobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			Cobros.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Tip_Ope     AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Mon_Net     AS  Cob_Ope, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Doc_Des     AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Num_Doc     AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Ban, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Caj, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpTemporalCobros, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Detalles_Cobros, ")
            loComandoSeleccionar.AppendLine(" 			Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE		Cobros.Documento                =   #tmpTemporalCobros.Documento ")
            loComandoSeleccionar.AppendLine("           AND Cobros.Documento            =   Detalles_Cobros.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Cli              =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Ven              =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 			AND Cobros.Cod_Mon              =   Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 			AND Detalles_Cobros.Cod_Cue     Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)
'me.mEscribirConsulta(loComandoSeleccionar.ToString)


            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCobros_Fechas_CCMV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Fechas_CCMV.ReportSource = loObjetoReporte

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
' JJD: 06/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
