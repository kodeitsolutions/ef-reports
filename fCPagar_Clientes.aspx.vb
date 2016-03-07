'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCPagar_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class fCPagar_Clientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Bru *(-1) Else Cuentas_Cobrar.Mon_Bru End) As Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Net *(-1) Else Cuentas_Cobrar.Mon_Net End) As Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End) As Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli          =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Mon      =   Monedas.Cod_Mon ")
			loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.status not in ('Anulado') ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Cobrar.Cod_Cli,  Cuentas_Cobrar.Cod_Tip, Cuentas_Cobrar.Documento ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
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

			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCPagar_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCPagar_Clientes.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS:  23/02/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplico el metodo carga de imagen
'-------------------------------------------------------------------------------------------'