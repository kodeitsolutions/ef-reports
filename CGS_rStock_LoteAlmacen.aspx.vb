'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rStock_LoteAlmacen"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rStock_LoteAlmacen
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Almacenes.Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones_Almacenes.Cod_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones_Almacenes.Cod_Art,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Uni1,")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine("		Secciones.Nom_Sec,")
            loComandoSeleccionar.AppendLine("		Renglones_Almacenes.Exi_Act1,")
            loComandoSeleccionar.AppendLine("		''  AS Cod_Lot")
            loComandoSeleccionar.AppendLine("FROM Renglones_Almacenes")
            loComandoSeleccionar.AppendLine("	    JOIN Almacenes ON Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	    JOIN Articulos ON Articulos. Cod_Art = Renglones_Almacenes.Cod_Art")
            loComandoSeleccionar.AppendLine("	    JOIN Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("	    JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("		    AND Departamentos.Cod_Dep = Secciones.Cod_Dep")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Usa_Lot = 0")
            loComandoSeleccionar.AppendLine("	AND Renglones_Almacenes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("   AND Almacenes.Cod_Alm BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Art BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Dep BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Sec BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Almacenes.Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones_Lotes.Cod_Alm,	")
            loComandoSeleccionar.AppendLine("		Renglones_Lotes.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Uni1,")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine("		Secciones.Cod_Sec, ")
            loComandoSeleccionar.AppendLine(" 		Renglones_Lotes.Exi_Act1, ")
            loComandoSeleccionar.AppendLine(" 		Renglones_Lotes.Cod_Lot 		")
            loComandoSeleccionar.AppendLine(" FROM Lotes ")
            loComandoSeleccionar.AppendLine("	    JOIN Renglones_Lotes ON Renglones_Lotes.Cod_Lot = Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("	    JOIN Almacenes ON Renglones_Lotes.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("	    JOIN Articulos ON Articulos.Cod_Art = Lotes.Cod_Art ")
            loComandoSeleccionar.AppendLine("	    JOIN Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("	    JOIN Secciones ON Secciones.Cod_Sec = Articulos.Cod_Sec")
            loComandoSeleccionar.AppendLine("		    AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Usa_Lot = 1")
            loComandoSeleccionar.AppendLine("	AND Renglones_Lotes.Exi_Act1 > 0")
            loComandoSeleccionar.AppendLine("   AND Almacenes.Cod_Alm BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Art BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Dep BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Sec BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro3Hasta)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rStock_LoteAlmacen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rStock_LoteAlmacen.ReportSource = loObjetoReporte

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
' CMS: 28/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'