'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProposiciones_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rProposiciones_Numeros
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT		Proposiciones.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine(" 				Prospectos.Nom_Pro, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Status, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Control, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM			Prospectos, ")
            loComandoSeleccionar.AppendLine(" 				Proposiciones, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores, ")
            loComandoSeleccionar.AppendLine(" 				Transportes, ")
            loComandoSeleccionar.AppendLine(" 				Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE			Proposiciones.Cod_Pro = Prospectos.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Pro between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Tra between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Proposiciones.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND Proposiciones.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("               AND Proposiciones.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY Proposiciones." & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProposiciones_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProposiciones_Numeros.ReportSource = loObjetoReporte


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
' MAT: 16/02/11: Codigo inicial
'-------------------------------------------------------------------------------------------'

