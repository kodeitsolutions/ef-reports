Imports System.Data
Partial Class TRF_rProveedores_dBasicos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde	VARCHAR(10) = ''")
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta	VARCHAR(10) = 'zzzzzz'")
            loComandoSeleccionar.AppendLine("DECLARE @lcCodZon_Desde	VARCHAR(10) = ''")
            loComandoSeleccionar.AppendLine("DECLARE @lcCodZon_Hasta	VARCHAR(10) = 'zzzzzz'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("		REPLACE(Proveedores.Rif,'-', '')	AS Rif, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Telefonos,")
            loComandoSeleccionar.AppendLine("		Proveedores.Dir_Fis,")
            loComandoSeleccionar.AppendLine("		Estados.Nom_Est")
            loComandoSeleccionar.AppendLine("FROM Proveedores ")
            loComandoSeleccionar.AppendLine("	JOIN Zonas ON Proveedores.Cod_Zon = Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine("	JOIN Estados ON Proveedores.Cod_Est = Estados.Cod_Est ")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine(" AND Proveedores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" AND Zonas.Cod_Zon BETWEEN @lcCodZon_Desde AND @lcCodZon_Hasta")
            loComandoSeleccionar.AppendLine(" AND Proveedores.Telefonos <> ''")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro, Nom_Pro")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rProveedores_dBasicos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rProveedores_dBasicos.ReportSource = loObjetoReporte


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
' MVP:  10/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  21/04/09: Estandarizacion de codigos y correccion de campo estatus
'-------------------------------------------------------------------------------------------'
' EAG:  28/09/15: Se acomodó el select, debido a que se generaba un error con una columna
'-------------------------------------------------------------------------------------------'