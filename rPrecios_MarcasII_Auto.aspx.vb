'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrecios_MarcasII_Auto"
'-------------------------------------------------------------------------------------------'
Partial Class rPrecios_MarcasII_Auto
    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("SELECT	Marcas.Nom_Mar,	")
        loComandoSeleccionar.AppendLine("		Articulos.Cod_Art						AS Codigo, ")
        loComandoSeleccionar.AppendLine("		Articulos.Nom_Art						AS Descripción,")
        loComandoSeleccionar.AppendLine("		Articulos.modelo						AS modelo,")
        loComandoSeleccionar.AppendLine("		Clases_Articulos.nom_cla							AS Observacion,")
        loComandoSeleccionar.AppendLine("		Articulos.Exi_Act1-Articulos.Exi_Ped1	AS Stock,")
        Select Case lcParametro9Desde
            Case "Precio1"
                loComandoSeleccionar.AppendLine("           Articulos.Precio1 AS Precio")
            Case "Precio2"
                loComandoSeleccionar.AppendLine("           Articulos.Precio2 AS Precio")
            Case "Precio3"
                loComandoSeleccionar.AppendLine("           Articulos.Precio3 AS Precio")
            Case "Precio4"
                loComandoSeleccionar.AppendLine("           Articulos.Precio4 AS Precio")
            Case "Precio5"
                loComandoSeleccionar.AppendLine("           Articulos.Precio5 AS Precio")
        End Select
        loComandoSeleccionar.AppendLine("FROM Articulos")
        loComandoSeleccionar.AppendLine("       JOIN	Departamentos	ON  Articulos.Cod_Dep       =   Departamentos.Cod_Dep ")
        loComandoSeleccionar.AppendLine("       JOIN	Secciones		ON	Articulos.Cod_Sec       =   Secciones.Cod_Sec")
        loComandoSeleccionar.AppendLine("           And Departamentos.Cod_Dep   =   Secciones.Cod_Dep  ")
        loComandoSeleccionar.AppendLine("       JOIN	Marcas			ON	Articulos.Cod_Mar       =   Marcas.Cod_Mar")
        loComandoSeleccionar.AppendLine(" 	  JOIN	Tipos_Articulos	ON	Articulos.Cod_Tip       =   Tipos_Articulos.Cod_Tip")
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       =   Tipos_Articulos.Cod_Tip ")
        loComandoSeleccionar.AppendLine(" 	  JOIN	Clases_Articulos ON	Articulos.Cod_Cla       =   Clases_Articulos.Cod_Cla ")
        loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Art       Between " & lcParametro0Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
        loComandoSeleccionar.AppendLine("           And Articulos.Cod_Ubi    Between " & lcParametro7Desde)
        loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
        If lcParametro10Desde = "Si" Then
            loComandoSeleccionar.AppendLine("       And Articulos.Exi_Act1-Articulos.Exi_Ped1 > 0")
        End If
        loComandoSeleccionar.AppendLine("ORDER BY   Marcas.nom_mar, " & lcParametro8Desde)

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
        'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
        If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                      "No se Encontraron Registros para los Parámetros Especificados. ", _
                                       vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                       "350px", _
                                       "200px")
        End If


        loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPrecios_MarcasII_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)

        Me.mFormatearCamposReporte(loObjetoReporte)

        Me.crvrPrecios_MarcasII_Auto.ReportSource = loObjetoReporte


        'Catch loExcepcion As Exception

        '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '                  "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '                   vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '                   "auto", _
        '                   "auto")

        'End Try

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
' EAG: 24/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
' EAG: 30/09/15: Cambió el campo que se muestra en la columna Observación paso de
'                Articulos.clase a Clases_Articulos.nom_cla                             
'-------------------------------------------------------------------------------------------'
