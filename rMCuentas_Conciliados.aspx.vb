'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMCajas_Fechas"
'-------------------------------------------------------------------------------------------'


Partial Class rMCuentas_Conciliados

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

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            ' loComandoSeleccionar.AppendLine("           Tipos_Movimientos.Nom_Tip,")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cuentas ")
            'loComandoSeleccionar.AppendLine(" JOIN  Tipos_Movimientos ON  Movimientos_Cuentas.Cod_Tip    =   Tipos_Movimientos.Cod_Tip")
            loComandoSeleccionar.AppendLine(" JOIN  Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine(" JOIN  Conceptos ON Movimientos_Cuentas.Cod_Con     =   Conceptos.Cod_Con")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Doc_Cil     <>   '0' ")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Documento   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue     Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Mon     Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Status      IN  (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Tip   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Rev     Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Suc     Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY    Movimientos_Cuentas.Cod_Cue, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMCuentas_Conciliados", laDatosReporte)

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
' CMS:  04/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  04/07/09: Se agregaron los siguientes filtros: Tipos de movimientos, Revisión, Sucursal
'-------------------------------------------------------------------------------------------'
' MAT:  29/09/11: Ajuste del Select y de la vista de diseño
'-------------------------------------------------------------------------------------------'