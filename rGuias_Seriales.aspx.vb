'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rGuias_Seriales"
'-------------------------------------------------------------------------------------------'
Partial Class rGuias_Seriales
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
           


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("SELECT	Seriales.Tip_Sal As Origen,")
            loComandoSeleccionar.AppendLine("          Seriales.doc_sal,")
            loComandoSeleccionar.AppendLine("          Guias.Documento,")
            loComandoSeleccionar.AppendLine("          Guias.Cod_Cli,")
            loComandoSeleccionar.AppendLine("          Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine("          Clientes.Rif,")
            loComandoSeleccionar.AppendLine("          Clientes.Dir_Fis,")
            loComandoSeleccionar.AppendLine("          Renglones_Guias.Cod_Art,")
            loComandoSeleccionar.AppendLine("          Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("          Articulos.Modelo AS Control,")
            loComandoSeleccionar.AppendLine("          Guias.Cod_Ven,")
            loComandoSeleccionar.AppendLine("          Vendedores.Nom_Ven,")
            loComandoSeleccionar.AppendLine("          Guias.Cod_Tra,")
            loComandoSeleccionar.AppendLine("          Transportes.Nom_Tra,")
            loComandoSeleccionar.AppendLine("          Clientes.Telefonos,")
            loComandoSeleccionar.AppendLine("          Guias.Fec_Ini,")
            loComandoSeleccionar.AppendLine("          Renglones_Guias.Renglon,")
            loComandoSeleccionar.AppendLine("          Renglones_Guias.Can_Art1,")
            loComandoSeleccionar.AppendLine("          ISNULL(Seriales.Serial,'') AS Serial")
            loComandoSeleccionar.AppendLine("FROM Guias")
            loComandoSeleccionar.AppendLine("JOIN Renglones_Guias ON Renglones_Guias.Documento = Guias.Documento")
            loComandoSeleccionar.AppendLine("JOIN Articulos ON Articulos.Cod_Art = Renglones_Guias.Cod_Art")
            loComandoSeleccionar.AppendLine("JOIN Clientes ON Clientes.Cod_Cli = Guias.Cod_Cli")
            loComandoSeleccionar.AppendLine("JOIN Vendedores ON Vendedores.Cod_Ven = Guias.Cod_Ven")
            loComandoSeleccionar.AppendLine("JOIN Transportes ON Transportes.Cod_Tra = Guias.Cod_Tra")
            loComandoSeleccionar.AppendLine("JOIN Seriales ON ((Seriales.Tip_Sal = Renglones_Guias.Tip_Ori")
            loComandoSeleccionar.AppendLine("                   AND Seriales.Doc_Sal = Renglones_Guias.Doc_ori")
            loComandoSeleccionar.AppendLine("                   AND Seriales.Ren_Sal = Renglones_Guias.Renglon)")
            loComandoSeleccionar.AppendLine("               OR (Seriales.Doc_Sal = Renglones_Guias.Documento")
            loComandoSeleccionar.AppendLine("                   AND	 Seriales.tip_sal	=	'Guias'")
            loComandoSeleccionar.AppendLine("                   AND Seriales.Ren_Sal = Renglones_Guias.Renglon))")
            loComandoSeleccionar.AppendLine("WHERE  Guias.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Guias.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Guias.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Guias.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND Guias.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Guias.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGuias_Seriales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrGuias_Seriales.ReportSource = loObjetoReporte

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
' MAT: 07/10/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
