'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCompras_AgrupadoArticulos"
'-------------------------------------------------------------------------------------------'
Partial Class rCompras_AgrupadoArticulos

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
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
			Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            
            Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine(" SELECT		Renglones_Compras.Cod_Art,")
			loComandoSeleccionar.AppendLine(" 				Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine(" 				Renglones_Compras.Cod_Uni,")
			loComandoSeleccionar.AppendLine(" 				SUM(Renglones_Compras.can_art1)	AS Can_Art1,")
			loComandoSeleccionar.AppendLine(" 				SUM(Renglones_Compras.Mon_Bru)	AS Mon_Bru,")
			loComandoSeleccionar.AppendLine(" 				SUM(Renglones_Compras.Mon_Imp1)	AS Mon_Imp1,")
			loComandoSeleccionar.AppendLine(" 				SUM(Renglones_Compras.Mon_Des)	AS Mon_Des")
			loComandoSeleccionar.AppendLine(" FROM			Compras, ")
			loComandoSeleccionar.AppendLine("				Renglones_Compras, ")
			loComandoSeleccionar.AppendLine("				Articulos	")
			loComandoSeleccionar.AppendLine(" WHERE			Compras.Documento = Renglones_Compras.Documento")
			loComandoSeleccionar.AppendLine("				 AND Renglones_Compras.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine("          		 AND Renglones_Compras.Cod_Art       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Compras.Fec_Ini                 BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Compras.Cod_Pro				 BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Compras.Cod_Ven                 BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Articulos.Cod_Dep               BETWEEN" & lcParametro4Desde )
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Articulos.Cod_Sec               BETWEEN" & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Articulos.Cod_Mar               BETWEEN" & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Renglones_Compras.Cod_Alm       BETWEEN " & lcParametro7Desde )
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("          		 AND Compras.Status                  IN (" & lcParametro8Desde  & ")" )
            
            loComandoSeleccionar.AppendLine("          		 AND Compras.Cod_rev                 BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("          		 AND " & lcParametro9Hasta)
            
			loComandoSeleccionar.AppendLine(" GROUP BY		Renglones_Compras.Cod_Art,")
			loComandoSeleccionar.AppendLine("				Articulos.Nom_Art,")
			loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Uni")
            'loComandoSeleccionar.AppendLine(" ORDER BY		Renglones_Compras.Cod_Art,")
            'loComandoSeleccionar.AppendLine("				Articulos.Nom_Art") 
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCompras_AgrupadoArticulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_AgrupadoArticulos.ReportSource = loObjetoReporte

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
' GCR: 11/03/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'