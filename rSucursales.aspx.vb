'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSucursales"
'-------------------------------------------------------------------------------------------'
Partial Class rSucursales
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()
			loComandoSeleccionar.AppendLine("  SELECT	Cod_Suc, ")
            loComandoSeleccionar.AppendLine("	        Nom_Suc, ")
            loComandoSeleccionar.AppendLine("           case Status ")
            loComandoSeleccionar.AppendLine("	            when  'A' then 'Activa' ")
            loComandoSeleccionar.AppendLine("	            when  'I' then 'Inactiva' ")
            loComandoSeleccionar.AppendLine("	            when  'S' then 'Suspendida' ")
            loComandoSeleccionar.AppendLine("           end as Status, ")
            loComandoSeleccionar.AppendLine("	        Contacto, ")
            loComandoSeleccionar.AppendLine("	        Fax, ")
            loComandoSeleccionar.AppendLine("	        Telefonos, ")
            loComandoSeleccionar.AppendLine("	        Direccion, ")
            loComandoSeleccionar.AppendLine("	        Correo ")
            loComandoSeleccionar.AppendLine(" FROM      Sucursales ")
            loComandoSeleccionar.AppendLine(" WHERE     ")
            loComandoSeleccionar.AppendLine(" 			Cod_Suc BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Status IN ( " & lcParametro1Desde & " ) ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Nom_Suc ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSucursales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSucursales.ReportSource = loObjetoReporte

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
' CMS: 22/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
