'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_Direcciones"
'-------------------------------------------------------------------------------------------'

Partial Class rClientes_Direcciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT			 Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("                Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Clientes.Dir_Fis AS VARCHAR(500)))) AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Clientes.Dir_Ent AS VARCHAR(500)))) AS Dir_Ent, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Clientes.Dir_Exa AS VARCHAR(500)))) AS Dir_Exa, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Clientes.Dir_Otr AS VARCHAR(500)))) AS Dir_Otr, ")
            loComandoSeleccionar.AppendLine("                Clientes.Dir_Ent, ")
            loComandoSeleccionar.AppendLine("                Clientes.Dir_Exa, ")
            loComandoSeleccionar.AppendLine("                Clientes.Dir_Otr, ")
            loComandoSeleccionar.AppendLine("                Clientes.Correo,")
            loComandoSeleccionar.AppendLine("                Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("                Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("                Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("                Zonas.Nom_Zon, ")
            loComandoSeleccionar.AppendLine("                Ciudades.Nom_Ciu, ")
            loComandoSeleccionar.AppendLine("                Estados.Nom_Est ")
            loComandoSeleccionar.AppendLine("FROM			 Clientes, ")
            loComandoSeleccionar.AppendLine("                Zonas, ")
            loComandoSeleccionar.AppendLine("                Ciudades, ")
            loComandoSeleccionar.AppendLine("                Estados, ")
            loComandoSeleccionar.AppendLine("                Vendedores ")
            loComandoSeleccionar.AppendLine("WHERE			 Clientes.Cod_Zon = Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Ciu = Ciudades.Cod_Ciu ")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Est = Estados.Cod_Est ")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Cli between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Cla between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Ven between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		 Clientes.Cod_Cli, Clientes.Nom_Cli")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_Direcciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_Direcciones.ReportSource = loObjetoReporte


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
' CMS:  04/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
