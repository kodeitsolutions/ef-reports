'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rMovimientosBancos"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rMovimientosBancos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @Fecha	AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @CodCue_Ini AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @CodCue_Fin AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10) = 0")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cuentas_Bancarias.Cod_Cue               AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       Cuentas_Bancarias.Nom_Cue               AS Nom_Cue, ")
            'loComandoSeleccionar.AppendLine("		Cuentas_Bancarias.Sal_Con +")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)")
            loComandoSeleccionar.AppendLine("				  FROM Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("				  WHERE Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("                   AND Movimientos_Cuentas.Fec_Ini < @Fecha")
            'loComandoSeleccionar.AppendLine("                   AND Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("					AND Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue),@lnCero)	AS Saldo_Anterior,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT SUM(Movimientos_Cuentas.Mon_Deb) ")
            loComandoSeleccionar.AppendLine("				  FROM Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("				  WHERE Movimientos_Cuentas.Fec_Ini >= @Fecha AND Movimientos_Cuentas.Fec_Ini < DATEADD(dd, DATEDIFF(dd, 0, @Fecha) + 1, 0)")
            loComandoSeleccionar.AppendLine("				    AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("					AND Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue),@lnCero)	AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT SUM(Movimientos_Cuentas.Mon_Hab) ")
            loComandoSeleccionar.AppendLine("				  FROM Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine("				  WHERE Movimientos_Cuentas.Fec_Ini >= @Fecha AND Movimientos_Cuentas.Fec_Ini < DATEADD(dd, DATEDIFF(dd, 0, @Fecha) + 1, 0)")
            loComandoSeleccionar.AppendLine("					AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("					AND Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue),@lnCero)	AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@Fecha									AS Fecha")
            loComandoSeleccionar.AppendLine("FROM   Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Bancarias.Cod_Cue BETWEEN @CodCue_Ini AND @CodCue_Fin")
            loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Bancarias.Cod_Cue, Cuentas_Bancarias.Nom_Cue, Cuentas_Bancarias.Fec_Cil, Cuentas_Bancarias.Sal_Con ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rMovimientosBancos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rMovimientosBancos.ReportSource = loObjetoReporte


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
' CMS: 22/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 31/07/09: Filtro "Revision:", Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'