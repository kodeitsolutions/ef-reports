'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rContadores"
'-------------------------------------------------------------------------------------------'
Partial Class rContadores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Contadores.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Contadores.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Contadores.Status, ")
            loComandoSeleccionar.AppendLine("           Contadores.Valor, ")
            loComandoSeleccionar.AppendLine("           Contadores.Cod_Suc, ")
            loComandoSeleccionar.AppendLine("           Sucursales.Nom_Suc, ")
            loComandoSeleccionar.AppendLine("           (Case When Contadores.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Contadores ")
            loComandoSeleccionar.AppendLine(" FROM      Contadores, ")
            loComandoSeleccionar.AppendLine("           Sucursales ")
            loComandoSeleccionar.AppendLine(" WHERE     Contadores.Cod_Suc          =   Sucursales.Cod_Suc ")
            loComandoSeleccionar.AppendLine("           And Contadores.Cod_Con      Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Sucursales.Cod_Suc      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Sucursales.Cod_Suc, Contadores.Cod_Con ")
            loComandoSeleccionar.AppendLine("ORDER BY    Sucursales.Cod_Suc, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rContadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrContadores.ReportSource = loObjetoReporte

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
' MJP: 09/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' JJD: 10/01/09 : Ajuste al reporte
'-------------------------------------------------------------------------------------------'
' CMS: 03/07/09 : Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'