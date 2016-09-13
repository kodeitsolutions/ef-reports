Imports System.Data
Partial Class rUsuarios_Globales
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loComandoSeleccionar As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Try

            loComandoSeleccionar.AppendLine("  SELECT	Cod_Usu, ")
            loComandoSeleccionar.AppendLine("  Nom_Usu, ")
            loComandoSeleccionar.AppendLine("  Status, ")
            loComandoSeleccionar.AppendLine("  (Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Usuarios ")
            loComandoSeleccionar.AppendLine("  FROM	 Usuarios ")
            loComandoSeleccionar.AppendLine("  WHERE Cod_Usu between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("  And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("  AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("  And Status In (" & lcParametro1Desde & " )")
            loComandoSeleccionar.AppendLine("  ORDER BY Cod_Usu, Nom_Usu")


            Dim loServicios As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rUsuarios_Globales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrUsuarios_Globales.ReportSource = loObjetoReporte

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
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JJD: 23/01/10: Ajustes al reporte
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'