Imports System.Data
Partial Class rEmpresas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
		    Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Empresas.Cod_Emp, ")
            loComandoSeleccionar.AppendLine("           Empresas.Nom_Emp, ")
            loComandoSeleccionar.AppendLine("           (Case When Empresas.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Empresas ")
            loComandoSeleccionar.AppendLine(" FROM      Empresas ")
            loComandoSeleccionar.AppendLine(" WHERE     Empresas.Cod_Emp     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Empresas.Status  IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		    AND Empresas.Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
           	loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
          

            Dim loServicios As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEmpresas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEmpresas.ReportSource = loObjetoReporte

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
' MJP: 09/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08: Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08: Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 21/03/09: Normalizacion del codigo 
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'