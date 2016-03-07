'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rInventario_Fisico "
'-------------------------------------------------------------------------------------------'
Partial Class rInventario_Fisico 
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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine("			Departamentos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("			Secciones.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("			Tipos_Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Clases_Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("			Marcas.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("FROM		Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Departamentos, ")
            loComandoSeleccionar.AppendLine(" 			Secciones, ")
            loComandoSeleccionar.AppendLine(" 			Marcas, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Clases_Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Proveedores ")
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("           AND Secciones.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.status IN (" & lcParametro1Desde & ")")

            loComandoSeleccionar.AppendLine("           AND Departamentos.Cod_Dep between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Secciones.Cod_Sec between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Marcas.Cod_Mar between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Tipos_Articulos.Cod_Tip between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Clases_Articulos.Cod_Cla between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Pro between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Ubi    Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY Articulos.Cod_Art, Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)
            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rInventario_Fisico", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrInventario_Fisico.ReportSource = loObjetoReporte


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
' MVP:  04/08/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP:  25/04/09 : Agreagar combo estatus y estandarizacion de codigo
'-------------------------------------------------------------------------------------------'
' CMS:  15/07/09 : Metodo de ordenamiento, Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'
' CMS:  24/08/09: Se agrego la union Secciones.Cod_Dep = Departamentos.Cod_Dep
'-------------------------------------------------------------------------------------------'