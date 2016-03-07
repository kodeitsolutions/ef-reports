'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rListasEmail_CorreosMarketing"
'-------------------------------------------------------------------------------------------'
Partial Class rListasEmail_CorreosMarketing
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))



        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try
            loConsulta.AppendLine("SELECT   Listas_Marketing.Cod_Lis AS Codigo,")
            loConsulta.AppendLine("         Listas_Marketing.Nom_Lis AS Nombre,")
            loConsulta.AppendLine("         (CASE WHEN Listas_Marketing.Status = 'A' THEN 'Activo' ELSE 'Inactivo' END) AS Status,")
            loConsulta.AppendLine("         Listas_Marketing.Tipo,")
            loConsulta.AppendLine("         Listas_Marketing.Clase,")
            loConsulta.AppendLine("         COALESCE(Renglones_lMarketing.Renglon,0) as RenglonCorreoLista,")
            loConsulta.AppendLine("         COALESCE( Renglones_lMarketing.Correo , '') as CorreoLista,")
            loConsulta.AppendLine("         COALESCE( (CASE  COALESCE(Renglones_lMarketing.Status,'') WHEN '' THEN NULL  WHEN 'A' THEN 'Activo' ELSE 'Inactivo' END), '') AS StatusCorreoLista")
            loConsulta.AppendLine("FROM	    Listas_Marketing")
            loConsulta.AppendLine("LEFT JOIN	    Renglones_lMarketing ON Listas_Marketing.Adicional='ListasEmail' AND Renglones_lMarketing.Cod_Lis = Listas_Marketing.Cod_Lis ")
            loConsulta.AppendLine("         AND  Listas_Marketing.Adicional= Renglones_lMarketing.Adicional ")

            loConsulta.AppendLine("WHERE      Listas_Marketing.Cod_Lis BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND      Listas_Marketing.Status IN (" & lcParametro1Desde & " )")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rListasEmail_CorreosMarketing", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrListasEmail_CorreosMarketing.ReportSource = loObjetoReporte

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
' JAC : 29/07/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'

