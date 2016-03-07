'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSecciones"
'-------------------------------------------------------------------------------------------'
Partial Class rSecciones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Secciones.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("           Secciones.Nom_Sec, ")
            loComandoSeleccionar.AppendLine("           Secciones.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("           Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine("           Secciones.Status, ")
            loComandoSeleccionar.AppendLine("           Case When Secciones.Status = 'A' Then 'Activo' Else 'Inactivo' End as Status_Secciones ")
            loComandoSeleccionar.AppendLine(" FROM      Secciones, ")
            loComandoSeleccionar.AppendLine("           Departamentos ")
            loComandoSeleccionar.AppendLine(" WHERE     Secciones.Cod_Dep       =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           And Secciones.Cod_Sec   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Secciones.Status    IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Secciones.Cod_Dep   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Secciones.Cod_Dep, ")
            'loComandoSeleccionar.AppendLine("           Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSecciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSecciones.ReportSource = loObjetoReporte

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
' MJP: 07/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08: Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08: Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 21/03/09: normalizacion del codigo. Inclusion del departamento y el status.
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'