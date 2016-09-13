Imports System.Data
Partial Class rUsuarios_Detallado

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cod_Usu, ")
            loComandoSeleccionar.AppendLine("           Nom_Usu, ")
            loComandoSeleccionar.AppendLine("           Tipo, ")
            loComandoSeleccionar.AppendLine("           Nivel, ")
            loComandoSeleccionar.AppendLine("           Sep_Fec, ")
            loComandoSeleccionar.AppendLine("           For_Fec, ")
            loComandoSeleccionar.AppendLine("           Sep_Mil, ")
            loComandoSeleccionar.AppendLine("           Sep_Dec, ")
            loComandoSeleccionar.AppendLine("           (Case Sexo When 'M' then 'Masculino' When 'F' Then 'Femenino' Else 'No Aplica' End) as Sexo, ")
            loComandoSeleccionar.AppendLine("           (Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status ")
            loComandoSeleccionar.AppendLine(" FROM      Usuarios ")
            loComandoSeleccionar.AppendLine(" WHERE     Cod_Usu      Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		    AND Status   IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Usu ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rUsuarios_Detallado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrUsuarios_Detallado.ReportSource = loObjetoReporte

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
' GCR: 01/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 14/07/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'