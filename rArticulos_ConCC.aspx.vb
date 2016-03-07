'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_ConCC
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

            'Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            'Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            'Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            'Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            'Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Tabla temporal con los registros a listar")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpRegistros( Codigo VARCHAR(30) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Nombre VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Estatus VARCHAR(15) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Contable XML")
            loComandoSeleccionar.AppendLine("                            );")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpRegistros(Codigo, Nombre, Estatus, Contable)")
            loComandoSeleccionar.AppendLine("SELECT  Cod_Art, ")
            loComandoSeleccionar.AppendLine("        Nom_Art,")
            loComandoSeleccionar.AppendLine("        (CASE Status ")
            loComandoSeleccionar.AppendLine("            WHEN 'A' THEN 'Activo'  ")
            loComandoSeleccionar.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("            ELSE 'Suspendido'  ")
            loComandoSeleccionar.AppendLine("        END) AS Status,")
            loComandoSeleccionar.AppendLine("        Contable")
            loComandoSeleccionar.AppendLine("FROM    Articulos")
            loComandoSeleccionar.AppendLine("WHERE   Articulos.Cod_Art           Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("        And Articulos.Cod_Ubi       Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- En el SELECT final se expande el XML Contable para obtener las ")
            loComandoSeleccionar.AppendLine("-- Cuentas Contables, de Gastos y Centros de Costos de cada página del registro ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  CASE WHEN (LEN(Detalles.Cue_Con_Codigo) > '1' AND (LEN(Detalles.Cue_Con_Codigo) < '9' OR LEN(Detalles.Cue_Con_Codigo) > '9')) THEN '******' ELSE '' END	AS Asteriscos,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Codigo                                AS Codigo,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Nombre                                AS Nombre,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Estatus                               AS Estatus,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Numero, 1)                        AS Numero,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Pagina, '')                       AS Pagina,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Con_Codigo, '')               AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Contables.Nom_Cue, '')             AS Cue_Con_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Gas_Codigo, '')               AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Gastos.Nom_Gas, '')                AS Cue_Gas_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Codigo, '')               AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Centros_Costos.Nom_Cen, '')                AS Cen_Cos_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Porcentaje, 0)            AS Cen_Cos_Porcentaje ")
            loComandoSeleccionar.AppendLine("FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  Codigo,")
            loComandoSeleccionar.AppendLine("                        (Ficha.C.value('@n[1]', 'VARCHAR(MAX)')+1) AS Numero,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('@nombre[1]', 'VARCHAR(MAX)') AS Pagina,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_con[1]', 'VARCHAR(MAX)') AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_gas[1]', 'VARCHAR(MAX)') AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("                        Costos.C.value('@codigo[1]', 'VARCHAR(MAX)') AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("                        CAST(Costos.C.value('@porcentaje[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)) AS Cen_Cos_Porcentaje")
            loComandoSeleccionar.AppendLine("                FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("                    CROSS APPLY Contable.nodes('contable/ficha') AS Ficha(C)")
            loComandoSeleccionar.AppendLine("                    OUTER APPLY Contable.nodes('contable/ficha/centro_costo') AS Costos(C)")
            loComandoSeleccionar.AppendLine("            ) Detalles")
            loComandoSeleccionar.AppendLine("        ON  Detalles.Codigo = #tmpRegistros.Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Contables")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Contables.Cod_Cue = Detalles.Cue_Con_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Gastos")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Gastos.Cod_Gas = Detalles.Cue_Gas_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Centros_Costos")
            loComandoSeleccionar.AppendLine("        ON Centros_Costos.Cod_Cen = Detalles.Cen_Cos_Codigo")
            'loComandoSeleccionar.AppendLine("WHERE    Detalles.Cue_Con_Codigo <> '' ")
            'loComandoSeleccionar.AppendLine("   AND   LEN(Detalles.Cue_Con_Codigo) <> '9' ")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpRegistros.Codigo, COALESCE(Detalles.Numero, 1)")
            loComandoSeleccionar.AppendLine("")



            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------'
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            '-------------------------------------------------------------------------------------------'
            ' Declaracion de variables para calculos de impuestos
            '-------------------------------------------------------------------------------------------'
            'Dim lcTipoImpuesto As String = ""
            'Dim lnValorImpuesto As Decimal = 0D
            'For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1

            '    '-------------------------------------------------------------------------------------------'
            '    ' Calcula el valor del impuesto dependiendo del tipo
            '    '-------------------------------------------------------------------------------------------'
            '    lnValorImpuesto = cusAdministrativo.goImpuestos.mObtenerPorcentaje(laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Cod_Imp"), DateTime.Now(), 10, lcTipoImpuesto)

            '    Select Case lcTipoImpuesto

            '        Case "Porcentaje"

            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Mon_Imp") = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Precio") * lnValorImpuesto / 100D
            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Por_Imp") = lnValorImpuesto
            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("lcTipoImpuesto") = "Porcentaje"

            '        Case "Monto"

            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Mon_Imp") = lnValorImpuesto
            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Por_Imp") = lnValorImpuesto
            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("lcTipoImpuesto") = "Monto"


            '        Case Else

            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Mon_Imp") = 0D
            '            laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Por_Imp") = 0D

            '    End Select


            'Next lnNumeroFila

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_ConCC.ReportSource = loObjetoReporte


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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' JJD: 31/10/14: Codigo inicial. Adecuacion para mostrar las cuentas contables              '
'-------------------------------------------------------------------------------------------'
' JJD: 18/12/14: Inclusion del Len de la Cuenta Contable                                    '
'-------------------------------------------------------------------------------------------'
