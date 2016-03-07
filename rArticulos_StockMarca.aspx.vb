'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_StockMarca"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_StockMarca
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
            
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcExisiencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Uni1, ")
            'lcComandoSeleccionar.AppendLine("           Articulos.Exi_Act1, ")
            
            Select Case lcParametro10Desde
                Case "Actual"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Act1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Act1"
                Case "Comprometida"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Ped1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Ped1"
                Case "Cotizada"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Cot1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Cot1"
                Case "En_Produccion"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Pro1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Pro1"
                Case "Por_Llegar"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Por1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Por1"
                Case "Por_Despachar"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Des1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Des1"
                Case "Por_Distribuir"
                    lcComandoSeleccionar.AppendLine("	Articulos.Exi_Dis1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Dis1"
            End Select
            
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Dep, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Sec, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
            lcComandoSeleccionar.AppendLine("           Marcas.Nom_Mar ")
            lcComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            lcComandoSeleccionar.AppendLine("           Departamentos, ")
            lcComandoSeleccionar.AppendLine("           Secciones, ")
            lcComandoSeleccionar.AppendLine("           Marcas, ")
            lcComandoSeleccionar.AppendLine("           Tipos_Articulos, ")
            lcComandoSeleccionar.AppendLine("           Clases_Articulos ")
            lcComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Dep           =   Departamentos.Cod_Dep ")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       =   Secciones.Cod_Sec ")
            lcComandoSeleccionar.AppendLine("           And Departamentos.Cod_Dep   =   Secciones.Cod_Dep ")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       =   Marcas.Cod_Mar ")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       =   Tipos_Articulos.Cod_Tip ")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       =   Clases_Articulos.Cod_Cla ")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Art       Between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("           And Articulos.Cod_Ubi    Between " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("      	And Articulos.Cod_Pro between " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine(" 		And " & lcParametro8Hasta)
            
            Select Case lcParametro9Desde
                Case "Todos"
                    lcComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    lcComandoSeleccionar.AppendLine("     AND Articulos.Exi_Act1          =   0  ")
                Case "Mayor"							  
                    lcComandoSeleccionar.AppendLine("     AND Articulos.Exi_Act1          >   0  ")
                Case "Menor"							  
                    lcComandoSeleccionar.AppendLine("     AND Articulos.Exi_Act1          <   0  ")
                Case "Maximo"							  
                    lcComandoSeleccionar.AppendLine("     AND Articulos.Exi_Max           =   " & lcExisiencia & "  ")
                Case "Minimo"							  
                    lcComandoSeleccionar.AppendLine("     And Articulos.Exi_Min           =   " & lcExisiencia & "  ")
                Case "Pedido"
                    lcComandoSeleccionar.AppendLine("     And Articulos.Exi_pto           =   " & lcExisiencia & "  ")
            End Select

            'lcComandoSeleccionar.AppendLine(" ORDER BY  Articulos.Cod_Art, Articulos.Nom_Art ")
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString(), "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_StockMarca", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_StockMarca.ReportSource = loObjetoReporte


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
' MVP: 09/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' JFP: 09/10/08: Limitacion de la longitud del nombre y ajustes de forma de presentacion.
'-------------------------------------------------------------------------------------------'
' JJD: 02/02/09: Ajustes a la aplicacion del filtro Status desde un combo.
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'
' CMS:  21/09/09: Se agregaro los filtros tipo de existencia, nivel de existencia y proveedor
'-------------------------------------------------------------------------------------------'