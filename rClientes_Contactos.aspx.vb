'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_Contactos"
'-------------------------------------------------------------------------------------------'
Partial Class rClientes_Contactos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" Select ")
            loComandoSeleccionar.AppendLine(" 		Clientes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		Contactos.Nom_Con,")
            loComandoSeleccionar.AppendLine(" 		Contactos.Telefonos,")
            loComandoSeleccionar.AppendLine(" 		Contactos.Correo")
            loComandoSeleccionar.AppendLine("  FROM  Clientes ")
            loComandoSeleccionar.AppendLine("  JOIN Contactos ON (Clientes.Cod_Cli collate Modern_Spanish_CI_AS = Contactos.Cod_Reg collate Modern_Spanish_CI_AS) ")
            loComandoSeleccionar.AppendLine("  AND ('" & goEmpresa.pcCodigo & "Clientes' collate Modern_Spanish_CI_AS = Contactos.Tipo collate Modern_Spanish_CI_AS) ")
            loComandoSeleccionar.AppendLine("  WHERE ")
            loComandoSeleccionar.AppendLine("     Clientes.Cod_Cli Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("     AND Clientes.Status In (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("  ORDER BY Clientes.Cod_Cli," & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_Contactos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrClientes_Contactos.ReportSource = loObjetoReporte

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
' CMS:  24/08/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  01/12/10 : Ajuste del Select
'-------------------------------------------------------------------------------------------'
