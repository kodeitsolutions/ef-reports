'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Seriales"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Seriales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


         
            loComandoSeleccionar.AppendLine("  SELECT   	")
            loComandoSeleccionar.AppendLine("  			Articulos.Cod_Art,")
            loComandoSeleccionar.AppendLine("  			Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("  			Articulos.Cod_uni1,")
            loComandoSeleccionar.AppendLine("  			Seriales.Serial,")
            loComandoSeleccionar.AppendLine("  			Seriales.Alm_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Ent,")
            loComandoSeleccionar.AppendLine("  			Seriales.Alm_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Tip_Sal,")
            loComandoSeleccionar.AppendLine("  			Seriales.Doc_Sal ")
            loComandoSeleccionar.AppendLine("  FROM Seriales")
            loComandoSeleccionar.AppendLine("  JOIN Articulos ON Seriales.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Art   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Dep   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Sec   Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Tip   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Cla   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Mar   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               And Articulos.Cod_Pro   Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("               And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" WHERE		Seriales.Serial   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            Select Case lcParametro8Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("")
                Case "Entregados"
                    loComandoSeleccionar.AppendLine("   And Seriales.Doc_Sal NOT IN ('')")
                Case "Por_Entregar"
                    loComandoSeleccionar.AppendLine("   And Seriales.Doc_Sal IN ('')")
            End Select

            loComandoSeleccionar.AppendLine("           And Seriales.Alm_Ent   Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Seriales.Alm_Sal   Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY  Articulos.Cod_Art, " & lcOrdenamiento)
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Seriales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrArticulos_Seriales.ReportSource = loObjetoReporte

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
' CMS: 27/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 15/02/11: Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'