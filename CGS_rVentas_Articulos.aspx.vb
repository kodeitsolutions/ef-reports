'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rVentas_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rVentas_Articulos
    Inherits vis2Formularios.frmReporte

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


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT  Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("        Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("        SUBSTRING(Articulos.Contable, 141,12)  AS CC_Art,")
            loComandoSeleccionar.AppendLine("		 (SELECT Nom_Cue")
            loComandoSeleccionar.AppendLine("		  FROM Cuentas_Contables")
            loComandoSeleccionar.AppendLine("		  WHERE Cod_Cue=SUBSTRING(Articulos.Contable, 141,12) ")
            loComandoSeleccionar.AppendLine("		 )                                      AS CCNom_Art,")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Notas,  ")
            loComandoSeleccionar.AppendLine("        Articulos.Cod_Dep,")
            loComandoSeleccionar.AppendLine("        Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine("        Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("        Secciones.Nom_Sec,")
            loComandoSeleccionar.AppendLine("        SUBSTRING(secciones.contable, 141,12)  AS CC_Sec,")
            loComandoSeleccionar.AppendLine("		 (SELECT Nom_Cue")
            loComandoSeleccionar.AppendLine("		  FROM Cuentas_Contables")
            loComandoSeleccionar.AppendLine("		  WHERE cod_cue=SUBSTRING(secciones.contable, 141,12) ")
            loComandoSeleccionar.AppendLine("		 )                                      AS CCNom_Sec,")
            loComandoSeleccionar.AppendLine("		 Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("        Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("        Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("        Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Por_Des, ")
            loComandoSeleccionar.AppendLine("        Renglones_Facturas.Mon_Net")
            loComandoSeleccionar.AppendLine("FROM   Articulos")
            loComandoSeleccionar.AppendLine("       JOIN Renglones_Facturas ON  Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("       JOIN Facturas ON  Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("       JOIN Clientes  ON Facturas.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("       JOIN Almacenes ON Renglones_Facturas.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("       JOIN Marcas ON Articulos.Cod_Mar = Marcas.Cod_Mar")
            loComandoSeleccionar.AppendLine("       JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("       JOIN Departamentos ON Secciones.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" WHERE Renglones_Facturas.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Facturas.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Dep BETWEEN" & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        AND Articulos.Cod_Sec BETWEEN" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY Renglones_Facturas.Cod_Art ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rVentas_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rVentas_Articulos.ReportSource = loObjetoReporte

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
' JJD: 09/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS: 04/08/09: Secciones.Cod_Dep = Departamentos.Cod_Dep
'-------------------------------------------------------------------------------------------'