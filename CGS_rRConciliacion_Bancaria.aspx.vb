'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRConciliacion_Bancaria"
'-------------------------------------------------------------------------------------------'


Partial Class CGS_rRConciliacion_Bancaria

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''' OBTIENE EL SALDO INICIAL CONCILIADO ANTES DE LA FECHA INDICADA EN LOS PARAMETROS''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaInicio	VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaFin		VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaInicio		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaFin		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldSaldoEstado		DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET @lcCuentaInicio		= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET @lcCuentaFin			= " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET @ldFechaInicio			= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET @ldFechaFin			= " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET @ldSaldoEstado			= " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Movimientos_Cuentas.Cod_Cue			                                        AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("		Cuentas_Bancarias.Num_Cue			                                            AS Num_Cue,")
            loComandoSeleccionar.AppendLine("		Bancos.Nom_Ban						                                            AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("		SUM(Movimientos_Cuentas.Mon_Deb)	                                            AS Cargos_Con,")
            loComandoSeleccionar.AppendLine("		SUM(Movimientos_Cuentas.Mon_Hab)	                                            AS Abonos_Con,")
            loComandoSeleccionar.AppendLine("        SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)                 AS Sal_Ini_Con,")
            loComandoSeleccionar.AppendLine("        (SUM(Movimientos_Cuentas.Mon_Deb) ")
            loComandoSeleccionar.AppendLine("         - SUM(Movimientos_Cuentas.Mon_Hab) ")
            loComandoSeleccionar.AppendLine("         + SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)")
            loComandoSeleccionar.AppendLine("        )                                                                               AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)  ")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Fec_Ini < @ldFechaInicio")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0) AS Sal_Inicial,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb)                            ")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)  AS Cargos_NoCon,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Hab)                            ")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)  AS Abonos_NoCon,")
            loComandoSeleccionar.AppendLine("        ISNULL(")
            loComandoSeleccionar.AppendLine("            ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)  ")
            loComandoSeleccionar.AppendLine("            FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("            WHERE       Movimientos_Cuentas.Fec_Ini < @ldFechaInicio")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)")
            loComandoSeleccionar.AppendLine("            +")
            loComandoSeleccionar.AppendLine("            ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb)                            ")
            loComandoSeleccionar.AppendLine("            FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("            WHERE   Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)")
            loComandoSeleccionar.AppendLine("            -")
            loComandoSeleccionar.AppendLine("            ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Hab)                            ")
            loComandoSeleccionar.AppendLine("            FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("            WHERE   Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)")
            loComandoSeleccionar.AppendLine("            ,0)                                                                            AS Sal_Libro,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb) - SUM(Movimientos_Cuentas.Mon_Hab)")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)       AS Sal_Fin_NoCon,")
            loComandoSeleccionar.AppendLine("        @ldSaldoEstado                                                                     AS Saldo_Cuenta,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Deb)  ")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)   AS Sal_Debe,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT SUM(Movimientos_Cuentas.Mon_Hab)  ")
            loComandoSeleccionar.AppendLine("        FROM        Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("        WHERE       Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("            AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin),0)   AS Sal_Haber")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("    JOIN	Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine("    JOIN	Bancos ON Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil = 1 ")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue, Cuentas_Bancarias.Num_Cue, Bancos.Nom_Ban")
           


            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRConciliacion_Bancaria", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCuentas_Conciliados.ReportSource = loObjetoReporte

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
' JJD: 21/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 04/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 04/07/09: Se agregaron los siguientes filtros: Tipos de movimientos, Revisión, Sucursal
'-------------------------------------------------------------------------------------------'
' RJG: 23/11/11: Ajuste en el los filtros de fechas del SELECT.								'
'-------------------------------------------------------------------------------------------'
