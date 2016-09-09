Imports System.Data
Partial Class PAS_rTrabajadores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("DECLARE @CodTra_Fin	AS VARCHAR(15);")
            lcComandoSeleccionar.AppendLine("DECLARE @CodTra_Ini	AS VARCHAR(15);")
            lcComandoSeleccionar.AppendLine("DECLARE @Contrato_Ini	AS VARCHAR(2);")
            lcComandoSeleccionar.AppendLine("DECLARE @Contrato_Fin	AS VARCHAR(2);")
            lcComandoSeleccionar.AppendLine("DECLARE @Dep_Ini		AS VARCHAR(3);")
            lcComandoSeleccionar.AppendLine("DECLARE @Dep_Fin		AS VARCHAR(3);")
            lcComandoSeleccionar.AppendLine("DECLARE @Cargo_Ini		AS VARCHAR(10);")
            lcComandoSeleccionar.AppendLine("DECLARE @Cargo_Fin		AS VARCHAR(10);")
            lcComandoSeleccionar.AppendLine("DECLARE @Fecha_Ini		AS DATETIME;")
            lcComandoSeleccionar.AppendLine("DECLARE @Fecha_Fin		AS DATETIME;")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SET @CodTra_Ini = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("SET @CodTra_Fin = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("SET @Contrato_Ini = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("SET @Contrato_Fin = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("SET @Dep_Ini = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("SET @Dep_Fin = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("SET @Cargo_Ini = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("SET @Cargo_Fin = " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("SET @Fecha_Ini = " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("SET @Fecha_Fin = " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT  DISTINCT ")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Rif		AS Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra		AS Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car				AS Cargo,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Fec_Ini		AS Ingreso,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recibos.Mon_Net	AS Sueldo,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Status			AS Status,")
            lcComandoSeleccionar.AppendLine("		@Fecha_Ini 					AS Desde,")
            lcComandoSeleccionar.AppendLine("		@Fecha_Fin 					AS Hasta")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Contratos ON Trabajadores.Cod_Con = Contratos.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("WHERE Recibos.Fecha >= @Fecha_Ini AND Recibos.Fecha < DATEADD(dd, DATEDIFF(dd, 0, @Fecha_Fin) + 1, 0)")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Recibos.Cod_Con = 'Q024'")
            lcComandoSeleccionar.AppendLine("	AND Trabajadores.Cod_Tra BETWEEN @CodTra_Ini AND @CodTra_Fin")
            lcComandoSeleccionar.AppendLine("	AND Contratos.Cod_Con BETWEEN @Contrato_Ini AND @Contrato_Fin")
            lcComandoSeleccionar.AppendLine("	AND Departamentos_Nomina.Cod_Dep BETWEEN @Dep_Ini AND @Dep_Fin")
            lcComandoSeleccionar.AppendLine("	AND Cargos.Cod_Car BETWEEN @Cargo_Ini AND @Cargo_Fin")
            lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro1Desde & " )")
            lcComandoSeleccionar.AppendLine("ORDER BY Trabajadores.Rif, Trabajadores.Nom_Tra")

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("PAS_rTrabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvPAS_rTrabajadores.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
