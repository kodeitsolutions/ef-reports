'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUsuarios_Grupos1"
'-------------------------------------------------------------------------------------------'
Partial Class fUsuarios_Grupos1

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Usuarios.Cod_Usu AS Cod_Usu,")
            loComandoSeleccionar.AppendLine(" 		Usuarios.Nom_Usu AS Nom_Usu,")
            loComandoSeleccionar.AppendLine(" 		Grupos.Cod_Gru AS Cod_Gru,")
            loComandoSeleccionar.AppendLine(" 		Grupos.Nom_Gru AS Nom_Gru,")
            loComandoSeleccionar.AppendLine(" 		Grupos.Accesos AS Accesos,")
            loComandoSeleccionar.AppendLine(" 		'' AS Modulo,")
            loComandoSeleccionar.AppendLine(" 		'' AS Seccion,")
            loComandoSeleccionar.AppendLine(" 		'' AS Opcion,")
            loComandoSeleccionar.AppendLine(" 		'' AS Agregar,")
            loComandoSeleccionar.AppendLine(" 		'' AS Editar,")
            loComandoSeleccionar.AppendLine(" 		'' AS Buscar,")
            loComandoSeleccionar.AppendLine(" 		'' AS Eliminar,")
            loComandoSeleccionar.AppendLine(" 		'' AS Imprimir")
            loComandoSeleccionar.AppendLine(" FROM	Usuarios")
            loComandoSeleccionar.AppendLine(" JOIN	Renglones_gUsuarios ON Usuarios.Cod_Usu = Renglones_gUsuarios.Cod_Usu")
            loComandoSeleccionar.AppendLine(" JOIN	Grupos ON Renglones_gUsuarios.Cod_Gru = Grupos.Cod_Gru")
            loComandoSeleccionar.AppendLine(" WHERE	" & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            If laDatosReporte.Tables(0).Rows.Count > 0 Then

                'Creamos la tabla
                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Cod_Usu", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Usu", GetType(String))
                loColumna.MaxLength = 150
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Gru", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nom_Gru", GetType(String))
                loColumna.MaxLength = 150
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Accesos", GetType(String))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Modulo", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Seccion", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Opcion", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Agregar", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Editar", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Buscar", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Eliminar", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Imprimir", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                'Variables para leer la tabla y la estructura xml
                Dim loFila As DataRow
                Dim loNuevaFila As DataRow
                Dim lcXmlStr As String = ""
                Dim loXml As New System.Xml.XmlDocument()
                Dim lcStrAcciones As String = ""

                'Recorre cada renglon de la tabla
                For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)

                    'Leemos el campo accesos que contiene la informacion xml del grupo
                    lcXmlStr = loFila.Item("Accesos")
                    'Cargamos la informacion en el objeto xml
                    loXml.LoadXml(lcXmlStr)

                    'Recorremos los nodos del xml
                    'Recorremos los nodos referente a los modulos de cada grupo
                    For Each loXmlModulo As System.Xml.XmlNode In loXml.SelectNodes("sistemas/sistema/modulo")
                        'Recorremos los nodos referente a las secciones de cada modulo
                        For Each loXmlSeccion As System.Xml.XmlNode In loXmlModulo.SelectNodes("seccion")
                            'Recorremos los nodos referente a los formularios de cada seccion
                            For Each loXmlFormulario As System.Xml.XmlNode In loXmlSeccion.SelectNodes("formulario")
                                'Leemos el campo de acciones del formulario
                                lcStrAcciones = loXmlFormulario.Attributes("acciones").InnerText()
                                'Creamos y cargamos la nueva fila de la tabla
                                loNuevaFila = loTabla.NewRow()
                                loTabla.Rows.Add(loNuevaFila)

                                loNuevaFila.Item("Cod_Usu") = loFila.Item("Cod_Usu")
                                loNuevaFila.Item("Nom_Usu") = loFila.Item("Nom_Usu")
                                loNuevaFila.Item("Cod_Gru") = loFila.Item("Cod_Gru")
                                loNuevaFila.Item("Nom_Gru") = loFila.Item("Nom_Gru")
                                loNuevaFila.Item("Accesos") = loFila.Item("Accesos")
                                loNuevaFila.Item("Modulo") = loXmlModulo.Attributes("nombre").InnerText()
                                loNuevaFila.Item("Seccion") = loXmlSeccion.Attributes("nombre").InnerText()
                                loNuevaFila.Item("Opcion") = loXmlFormulario.Attributes("opcion").InnerText()
                                loNuevaFila.Item("Agregar") = lcStrAcciones.Contains("agregar").ToString
                                loNuevaFila.Item("Editar") = lcStrAcciones.Contains("editar").ToString
                                loNuevaFila.Item("Buscar") = lcStrAcciones.Contains("buscar").ToString
                                loNuevaFila.Item("Eliminar") = lcStrAcciones.Contains("eliminar").ToString
                                loNuevaFila.Item("Imprimir") = lcStrAcciones.Contains("imprimir").ToString

                                loTabla.AcceptChanges()
                            Next
                        Next
                    Next
                Next

                Dim loDatosFinal As New DataSet("curReportes")
                loDatosFinal.Tables.Add(loTabla)

                '--------------------------------------------------'
                ' Carga la imagen del logo en cusReportes          '
                '--------------------------------------------------'
                Me.mCargarLogoEmpresa(loDatosFinal.Tables(0), "LogoEmpresa")

                '-------------------------------------------------------------------------------------------------------
                ' Verificando si el select (tabla nº0) trae registros
                '-------------------------------------------------------------------------------------------------------

                loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUsuarios_Grupos1", loDatosFinal)
            Else
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                            "No se Encontraron Registros para los Parámetros Especificados. ", _
                                            vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                            "350px", _
                                            "200px")
                    loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("fUsuarios_Grupos1", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfUsuarios_Grupos1.ReportSource = loObjetoReporte

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
' Douglas Cortez: 18/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'
