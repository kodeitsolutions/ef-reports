'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCPagar_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fCPagar_Proveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Pagar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tip_Doc, ")
            'loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Bru *(-1) Else Cuentas_Pagar.Mon_Bru End) As Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Imp1, ")
            'loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Net *(-1) Else Cuentas_Pagar.Mon_Net End) As Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Sal *(-1) Else Cuentas_Pagar.Mon_Sal End) As Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM      Proveedores, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Cod_Pro          =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven      =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tra      =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Mon      =   Monedas.Cod_Mon ")
	    loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.status not in ('Anulado') ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Pagar.Cod_Pro,  Cuentas_Pagar.Cod_Tip, Cuentas_Pagar.Documento ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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

            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCPagar_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCPagar_Proveedores.ReportSource = loObjetoReporte

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
' CMS:  18/05/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'