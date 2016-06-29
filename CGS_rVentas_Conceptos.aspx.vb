'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rVentas_Conceptos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rVentas_Conceptos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            'Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            'Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            'Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            'Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Conceptos.Cod_Con,")
            loComandoSeleccionar.AppendLine("		Conceptos.Nom_Con,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(Conceptos.Contable, 141,12)	AS CC_Con,")
            loComandoSeleccionar.AppendLine("		ISNULL(( SELECT Nom_Cue")
            loComandoSeleccionar.AppendLine("		  FROM Cuentas_Contables")
            loComandoSeleccionar.AppendLine("		  WHERE Cod_Cue=SUBSTRING(Conceptos.Contable, 141,12) ")
            loComandoSeleccionar.AppendLine("		),'')									AS CCNom_Con,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Comentario,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Factura,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Control,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("FROM Conceptos")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Con = Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Cobrar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Con BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY Cuentas_Cobrar.Cod_Con ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rVentas_Conceptos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rVentas_Conceptos.ReportSource = loObjetoReporte

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
' JJD: 09/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS: 04/08/09: Secciones.Cod_Dep = Departamentos.Cod_Dep
'-------------------------------------------------------------------------------------------'