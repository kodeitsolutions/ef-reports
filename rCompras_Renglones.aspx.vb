'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCompras_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCompras_Renglones
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT			Compras.Documento, ")
            loComandoSeleccionar.AppendLine("				Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("				Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Renglon, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Precio1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Por_Des, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Des, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("				Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("				Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("				Formas_Pagos.Nom_For ")
            loComandoSeleccionar.AppendLine("FROM			Compras, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras, ")
            loComandoSeleccionar.AppendLine("				Articulos, ")
            loComandoSeleccionar.AppendLine("				Proveedores, ")
            loComandoSeleccionar.AppendLine("				Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("				Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE		Compras.Documento = Renglones_Compras.Documento ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Compras.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_For = Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Pro BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Rev BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Compras.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCompras_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_Renglones.ReportSource = loObjetoReporte


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
' MVP:  15/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  17/03/09: Estandarización de código y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' YJP:  14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS y AAP:  24/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'

