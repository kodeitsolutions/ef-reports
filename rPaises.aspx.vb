'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPaises"
'-------------------------------------------------------------------------------------------'
Partial Class rPaises

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Paises.Cod_Pai, ")
            loComandoSeleccionar.AppendLine("           Paises.Nom_Pai, ")
            loComandoSeleccionar.AppendLine("           Paises.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Monedas.Nom_Mon, ")
            loComandoSeleccionar.AppendLine("           (Case When Paises.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status ")
            loComandoSeleccionar.AppendLine(" FROM      Paises, Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE     Paises.Cod_Pai     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Paises.Status  IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Paises.Cod_Mon = Monedas.Cod_Mon")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Paises.Cod_Pai, ")
            'loComandoSeleccionar.AppendLine("           Paises.Nom_Pai")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPaises", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPaises.ReportSource = loObjetoReporte

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
' MJP:  08/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP:  11/07/08: Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP:  14/07/08: Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD:  02/02/09: Se agrego la lista del Status
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'