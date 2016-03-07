'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrecios_Almacenes_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rPrecios_Almacenes_Resumido
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT     Articulos.Cod_Dep				AS Cod_Dep, ")
            loComandoSeleccionar.AppendLine("			Departamentos.Nom_Dep			AS Nom_Dep, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Cod_Alm		AS Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Almacenes.Nom_Alm				AS Nom_Alm, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art				AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art				AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Exi_Act1	AS Exi_Act1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Exi_Ped1	AS Exi_Ped1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Web					AS Web, ")


            Select Case lcParametro8Desde
                Case "Si"
                    loComandoSeleccionar.AppendLine("			'Si'						AS Disponible,")
                Case "No"
                    loComandoSeleccionar.AppendLine("			'No'						AS Disponible,")
            End Select

            Select Case lcParametro9Desde
                Case "Precio1"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio1 				AS Precio")
                Case "Precio2"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio2 				AS Precio")
                Case "Precio3"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio3 				AS Precio")
                Case "Precio4"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio4 				AS Precio")
                Case "Precio5"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio5				AS Precio")
            End Select

            loComandoSeleccionar.AppendLine("FROM		Articulos ")
            loComandoSeleccionar.AppendLine("	JOIN 	Departamentos ON (Articulos.Cod_Dep = Departamentos.Cod_Dep )")
            loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Almacenes ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN 	Almacenes  ")
            loComandoSeleccionar.AppendLine("		ON	Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Renglones_Almacenes.Cod_Alm	BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Art			BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Sec       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar       BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Tip       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Cla       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Ubi		BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)


            If lcParametro10Desde = "Si" Then
                loComandoSeleccionar.AppendLine("			AND (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1) > 0")
            End If

            loComandoSeleccionar.AppendLine("ORDER BY		" & lcOrdenamiento)

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
			
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPrecios_Almacenes_Resumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPrecios_Almacenes_Resumido.ReportSource = loObjetoReporte


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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MAT: 30/09/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 26/01/12: Ajuste en filtro de artículos con inventario (inventario por almacen).		'
'-------------------------------------------------------------------------------------------'
