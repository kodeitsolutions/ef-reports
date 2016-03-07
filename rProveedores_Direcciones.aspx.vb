'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProveedores_Direcciones"
'-------------------------------------------------------------------------------------------'

Partial Class rProveedores_Direcciones
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


            loComandoSeleccionar.AppendLine("SELECT			 Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("                Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Proveedores.Dir_Fis AS VARCHAR(500)))) AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Proveedores.Dir_Ent AS VARCHAR(500)))) AS Dir_Ent, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Proveedores.Dir_Exa AS VARCHAR(500)))) AS Dir_Exa, ")
            loComandoSeleccionar.AppendLine("                RTRIM(LTRIM(CAST(Proveedores.Dir_Otr AS VARCHAR(500)))) AS Dir_Otr, ")
            loComandoSeleccionar.AppendLine("                Proveedores.Correo,")
            loComandoSeleccionar.AppendLine("                Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("                Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("                Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("                Zonas.Nom_Zon, ")
            loComandoSeleccionar.AppendLine("                Ciudades.Nom_Ciu, ")
            loComandoSeleccionar.AppendLine("                Estados.Nom_Est ")
            loComandoSeleccionar.AppendLine("FROM			 Proveedores, ")
            loComandoSeleccionar.AppendLine("                Zonas, ")
            loComandoSeleccionar.AppendLine("                Ciudades, ")
            loComandoSeleccionar.AppendLine("                Estados, ")
            loComandoSeleccionar.AppendLine("                Vendedores ")
            loComandoSeleccionar.AppendLine("WHERE			 Proveedores.Cod_Zon = Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Ciu = Ciudades.Cod_Ciu ")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Est = Estados.Cod_Est ")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Pro between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Cla between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Ven between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		 Proveedores.Cod_Pro, Proveedores.Nom_Pro")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_Direcciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_Direcciones.ReportSource = loObjetoReporte


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
' CMS:  06/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
