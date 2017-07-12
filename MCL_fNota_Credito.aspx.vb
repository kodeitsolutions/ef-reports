'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_fNota_Credito"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_fNota_Credito
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT   Cuentas_Cobrar.Cod_Cli						AS Cod_Cli, ")
            loConsulta.AppendLine("         Clientes.Nom_Cli							AS Nom_Cli, ")
            loConsulta.AppendLine("         Clientes.Rif								AS Rif, ")
            loConsulta.AppendLine("         Clientes.Dir_Fis							AS Dir_Fis, ")
            loConsulta.AppendLine("         Clientes.Telefonos							AS Telefonos, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Documento					AS Documento, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Fec_Ini						AS Fec_Ini, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Mon_Bru						AS Mon_Bru, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Mon_Imp1					    AS Mon_Imp1, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Por_Imp1					    AS Por_Imp1, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Mon_Net						AS Mon_Net, ")
            loConsulta.AppendLine("         Formas_Pagos.Nom_For						AS Nom_For, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Comentario					AS Comentario, ")
            loConsulta.AppendLine("         COALESCE(Renglones_Documentos.Cod_Art,'')   AS Cod_Art, ")
            loConsulta.AppendLine("         COALESCE(Renglones_Documentos.Notas,'')     AS Notas, ")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Cobrar.Notas = ''")
            loConsulta.AppendLine("              THEN Articulos.Nom_Art")
            loConsulta.AppendLine("              ELSE Cuentas_Cobrar.Notas")
            loConsulta.AppendLine("         END										    AS Nom_Art,")
            loConsulta.AppendLine("         COALESCE(Renglones_Documentos.Can_Art,0)    AS Can_Art1, ")
            loConsulta.AppendLine("         COALESCE(Renglones_Documentos.Precio1,0)    AS Precio1,")
            loConsulta.AppendLine("         COALESCE(Renglones_Documentos.Mon_Net,0)	AS Neto, ")
            loConsulta.AppendLine("         COALESCE(Campos_Extras.Memo1,'')			AS Observacion")
            loConsulta.AppendLine("FROM Cuentas_Cobrar ")
            loConsulta.AppendLine(" LEFT JOIN Renglones_Documentos ON  Cuentas_Cobrar.Documento = Renglones_Documentos.Documento")
            loConsulta.AppendLine("	    AND Renglones_Documentos.Cod_Tip = Cuentas_Cobrar.Cod_Tip")
            loConsulta.AppendLine("	    AND Renglones_Documentos.Origen = 'Ventas'")
            loConsulta.AppendLine(" JOIN Clientes ON  Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli   ")
            loConsulta.AppendLine(" JOIN Formas_Pagos ON  Cuentas_Cobrar.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine(" LEFT JOIN Articulos ON  Articulos.Cod_Art = Renglones_Documentos.Cod_Art")
            loConsulta.AppendLine(" LEFT JOIN Campos_Extras ON Campos_Extras.Cod_Reg = Cuentas_Cobrar.Documento    ")
            loConsulta.AppendLine("     AND Campos_Extras.Origen = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("     AND Campos_Extras.Clase = 'N/CR'")
            loConsulta.AppendLine("WHERE" & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("     AND Cuentas_Cobrar.Cod_Tip = 'N/CR'")

            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("MCL_fNota_Credito", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvMCL_fNota_Credito.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'

