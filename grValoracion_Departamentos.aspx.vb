'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grValoracion_Departamentos"
'-------------------------------------------------------------------------------------------'
Partial Class grValoracion_Departamentos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		SUM(Articulos.Exi_Act1 * Articulos.Cos_Ult1) AS Val_Dep, ")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Nom_Dep ")
            loComandoSeleccionar.AppendLine(" INTO  #temporal1 ")
            loComandoSeleccionar.AppendLine(" FROM	Articulos, Departamentos ")
            loComandoSeleccionar.AppendLine(" WHERE	Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" Group By Departamentos.Cod_Dep, Departamentos.Nom_Dep ")
            loComandoSeleccionar.AppendLine(" ORDER BY Departamentos.Nom_Dep ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		SUM(#temporal1.Val_Dep) AS Tot_Val")
            loComandoSeleccionar.AppendLine(" INTO #temporal2")
            loComandoSeleccionar.AppendLine(" FROM #temporal1")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN #temporal2.Tot_Val > 0 THEN ROUND((#temporal1.Val_Dep/(#temporal2.Tot_Val))*100,2)")
            loComandoSeleccionar.AppendLine(" 			ELSE 0")
            loComandoSeleccionar.AppendLine(" 		END	AS Por_Val,")
            loComandoSeleccionar.AppendLine(" 		#temporal1.Val_Dep, ")
            loComandoSeleccionar.AppendLine(" 		#temporal1.Cod_Dep, ")
            loComandoSeleccionar.AppendLine(" 		#temporal1.Nom_Dep ")
            loComandoSeleccionar.AppendLine(" FROM #temporal1,#temporal2 ")

            'condición para mostrar solo los departamentos cuyo valor sea superior a cero
            loComandoSeleccionar.AppendLine(" WHERE #temporal1.Val_Dep > 0 ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("grValoracion_Departamentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrValoracion_Departamentos.ReportSource = loObjetoReporte

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
' Douglas Cortez 20/04/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' Douglas Cortez 20/04/2010: Mejorar la consulta SQL para que a futuro se pueda filtrar 
'                           los departamentos cuando su valor sea cero o no
'-------------------------------------------------------------------------------------------'