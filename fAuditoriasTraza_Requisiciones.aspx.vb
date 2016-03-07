'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAuditoriasTraza_Requisiciones"
'-------------------------------------------------------------------------------------------'
Partial Class fAuditoriasTraza_Requisiciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
			loConsulta.AppendLine("SET @lcDocumento = (")
			loConsulta.AppendLine("    SELECT TOP 1 Documento ")
			loConsulta.AppendLine("    FROM Requisiciones")
			loConsulta.AppendLine("    WHERE "  & cusAplicacion.goFormatos.pcCondicionPrincipal & ");")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("CREATE TABLE #tmpTraza( Nivel INT,")
			loConsulta.AppendLine("                        Orden INT,")
			loConsulta.AppendLine("                        Tipo CHAR(20) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Documento CHAR(10) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Renglon INT,")
			loConsulta.AppendLine("                        Estatus CHAR(15) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Fec_Ini DATE,")
			loConsulta.AppendLine("                        Fec_Fin DATE,")
			loConsulta.AppendLine("                        Tip_Ori CHAR(20) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Doc_Ori CHAR(10) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Ren_Ori INT,")
			loConsulta.AppendLine("                        Aud_Cod_Usu CHAR(10) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Aud_Registro DATETIME,")
			loConsulta.AppendLine("                        Aud_Tipo CHAR(15) COLLATE DATABASE_DEFAULT, ")
			loConsulta.AppendLine("                        Aud_Tabla CHAR(30) COLLATE DATABASE_DEFAULT, ")
			loConsulta.AppendLine("                        Aud_Cod_Obj CHAR(100) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Aud_Notas VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
			loConsulta.AppendLine("                        Aud_Opcion CHAR(100) COLLATE DATABASE_DEFAULT, ")
			loConsulta.AppendLine("                        Aud_Accion CHAR(10) COLLATE DATABASE_DEFAULT, ")
			loConsulta.AppendLine("                        Aud_Equipo CHAR(30) COLLATE DATABASE_DEFAULT, ")
			loConsulta.AppendLine("                        Aud_Detalle VARCHAR(MAX) COLLATE DATABASE_DEFAULT DEFAULT ('<detalle/>')")
			loConsulta.AppendLine("                        );")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- Requisición")
			loConsulta.AppendLine("INSERT INTO #tmpTraza(  Nivel, Orden, Tipo, Documento, Renglon,")
			loConsulta.AppendLine("                        Estatus, Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori, Ren_Ori,")
			loConsulta.AppendLine("                        Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                        Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, Aud_Detalle)")
			loConsulta.AppendLine("SELECT      1                                   AS Nivel,")
			loConsulta.AppendLine("            1                                   AS Orden,")
			loConsulta.AppendLine("            'Requisiciones'                     AS Tipo,")
			loConsulta.AppendLine("            Renglones_Requisiciones.Documento   AS Documento,")
			loConsulta.AppendLine("            Renglones_Requisiciones.Renglon     AS Renglon,")
			loConsulta.AppendLine("            Requisiciones.Status                AS Estatus,")
			loConsulta.AppendLine("            Requisiciones.Fec_Ini               AS Fec_Ini,")
			loConsulta.AppendLine("            Requisiciones.Fec_Fin               AS Fec_Fin,")
			loConsulta.AppendLine("            ''                                  AS Tip_Ori,")
			loConsulta.AppendLine("            ''                                  AS Doc_Ori,")
			loConsulta.AppendLine("            0                                   AS Ren_Ori,")
			loConsulta.AppendLine("            Auditorias.Cod_Usu                  AS Aud_Cod_Usu,")
			loConsulta.AppendLine("            Auditorias.Registro                 AS Aud_Registro,")
			loConsulta.AppendLine("            Auditorias.Tipo                     AS Aud_Tipo,")
			loConsulta.AppendLine("            Auditorias.Tabla                    AS Aud_Tabla,")
			loConsulta.AppendLine("            Auditorias.Cod_Obj                  AS Aud_Cod_Obj,")
			loConsulta.AppendLine("            Auditorias.Notas                    AS Aud_Notas,")
			loConsulta.AppendLine("            Auditorias.Opcion                   AS Aud_Opcion,")
			loConsulta.AppendLine("            Auditorias.Accion                   AS Aud_Accion,")
			loConsulta.AppendLine("            Auditorias.Equipo                   AS Aud_Equipo,")
			loConsulta.AppendLine("            Auditorias.Detalle                  AS Aud_Detalle")
			loConsulta.AppendLine("FROM        Requisiciones")
			loConsulta.AppendLine("    JOIN    Renglones_Requisiciones")
			loConsulta.AppendLine("        ON  Renglones_Requisiciones.Documento = Requisiciones.Documento")
			loConsulta.AppendLine("    LEFT JOIN Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Requisiciones.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Requisiciones' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('RequisicionesInternas', 'Sin opción')")
			loConsulta.AppendLine("WHERE       Renglones_Requisiciones.Documento = @lcDocumento;")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- Presupuestos")
			loConsulta.AppendLine("INSERT INTO #tmpTraza(  Nivel, Orden, Tipo, Documento, Renglon,")
			loConsulta.AppendLine("                        Estatus, Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori, Ren_Ori,")
			loConsulta.AppendLine("                        Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                        Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, Aud_Detalle)")
			loConsulta.AppendLine("SELECT      (#tmpTraza.Nivel + 1)               AS Nivel,")
			loConsulta.AppendLine("            2                                   AS Orden,")
			loConsulta.AppendLine("            'Presupuestos'                      AS Tipo,")
			loConsulta.AppendLine("            Renglones_Presupuestos.Documento    AS Documento,")
			loConsulta.AppendLine("            Renglones_Presupuestos.Renglon      AS Renglon,")
			loConsulta.AppendLine("            Presupuestos.Status                 AS Estatus,")
			loConsulta.AppendLine("            Presupuestos.Fec_Ini                AS Fec_Ini,")
			loConsulta.AppendLine("            Presupuestos.Fec_Fin                AS Fec_Fin,")
			loConsulta.AppendLine("            Renglones_Presupuestos.Tip_Ori      AS Tip_Ori,")
			loConsulta.AppendLine("            Renglones_Presupuestos.Doc_Ori      AS Doc_Ori,")
			loConsulta.AppendLine("            Renglones_Presupuestos.Ren_Ori      AS Ren_Ori,")
			loConsulta.AppendLine("            Auditorias.Cod_Usu                  AS Aud_Cod_Usu,")
			loConsulta.AppendLine("            Auditorias.Registro                 AS Aud_Registro,")
			loConsulta.AppendLine("            Auditorias.Tipo                     AS Aud_Tipo,")
			loConsulta.AppendLine("            Auditorias.Tabla                    AS Aud_Tabla,")
			loConsulta.AppendLine("            Auditorias.Cod_Obj                  AS Aud_Cod_Obj,")
			loConsulta.AppendLine("            Auditorias.Notas                    AS Aud_Notas,")
			loConsulta.AppendLine("            Auditorias.Opcion                   AS Aud_Opcion,")
			loConsulta.AppendLine("            Auditorias.Accion                   AS Aud_Accion,")
			loConsulta.AppendLine("            Auditorias.Equipo                   AS Aud_Equipo,")
			loConsulta.AppendLine("            Auditorias.Detalle                  AS Aud_Detalle")
			loConsulta.AppendLine("FROM        #tmpTraza")
			loConsulta.AppendLine("    JOIN    Renglones_Presupuestos ")
			loConsulta.AppendLine("        ON  Renglones_Presupuestos.Tip_Ori = #tmpTraza.Tipo")
			loConsulta.AppendLine("        AND Renglones_Presupuestos.Doc_Ori = #tmpTraza.Documento")
			loConsulta.AppendLine("        AND Renglones_Presupuestos.Ren_Ori = #tmpTraza.Renglon")
			loConsulta.AppendLine("    JOIN    Presupuestos")
			loConsulta.AppendLine("        ON  Presupuestos.Documento = Renglones_Presupuestos.Documento")
			loConsulta.AppendLine("    LEFT JOIN Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Presupuestos.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Presupuestos' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('Presupuestos', 'Sin opción');")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- Ordenes de Compra")
			loConsulta.AppendLine("INSERT INTO #tmpTraza(  Nivel, Orden, Tipo, Documento, Renglon,")
			loConsulta.AppendLine("                        Estatus, Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori, Ren_Ori,")
			loConsulta.AppendLine("                        Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                        Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, Aud_Detalle)")
			loConsulta.AppendLine("SELECT      (#tmpTraza.Nivel + 1)               AS Nivel,")
			loConsulta.AppendLine("            3                                   AS Orden,")
			loConsulta.AppendLine("            'Ordenes_Compras'                   AS Tipo,")
			loConsulta.AppendLine("            Renglones_oCompras.Documento        AS Documento,")
			loConsulta.AppendLine("            Renglones_oCompras.Renglon          AS Renglon,")
			loConsulta.AppendLine("            Ordenes_Compras.Status              AS Estatus,")
			loConsulta.AppendLine("            Ordenes_Compras.Fec_Ini             AS Fec_Ini,")
			loConsulta.AppendLine("            Ordenes_Compras.Fec_Fin             AS Fec_Fin,")
			loConsulta.AppendLine("            Renglones_oCompras.Tip_Ori          AS Tip_Ori,")
			loConsulta.AppendLine("            Renglones_oCompras.Doc_Ori          AS Doc_Ori,")
			loConsulta.AppendLine("            Renglones_oCompras.Ren_Ori          AS Ren_Ori,")
			loConsulta.AppendLine("            Auditorias.Cod_Usu                  AS Aud_Cod_Usu,")
			loConsulta.AppendLine("            Auditorias.Registro                 AS Aud_Registro,")
			loConsulta.AppendLine("            Auditorias.Tipo                     AS Aud_Tipo,")
			loConsulta.AppendLine("            Auditorias.Tabla                    AS Aud_Tabla,")
			loConsulta.AppendLine("            Auditorias.Cod_Obj                  AS Aud_Cod_Obj,")
			loConsulta.AppendLine("            Auditorias.Notas                    AS Aud_Notas,")
			loConsulta.AppendLine("            Auditorias.Opcion                   AS Aud_Opcion,")
			loConsulta.AppendLine("            Auditorias.Accion                   AS Aud_Accion,")
			loConsulta.AppendLine("            Auditorias.Equipo                   AS Aud_Equipo,")
			loConsulta.AppendLine("            Auditorias.Detalle                  AS Aud_Detalle")
			loConsulta.AppendLine("FROM        #tmpTraza")
			loConsulta.AppendLine("    JOIN    Renglones_oCompras ")
			loConsulta.AppendLine("        ON  Renglones_oCompras.Tip_Ori = #tmpTraza.Tipo")
			loConsulta.AppendLine("        AND Renglones_oCompras.Doc_Ori = #tmpTraza.Documento")
			loConsulta.AppendLine("        AND Renglones_oCompras.Ren_Ori = #tmpTraza.Renglon")
			loConsulta.AppendLine("    JOIN    Ordenes_Compras")
			loConsulta.AppendLine("        ON  Ordenes_Compras.Documento = Renglones_oCompras.Documento")
			loConsulta.AppendLine("    LEFT JOIN Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Ordenes_Compras.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Ordenes_Compras' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('OrdenesCompra', 'Sin opción');")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- Notas de Recepción")
			loConsulta.AppendLine("INSERT INTO #tmpTraza(  Nivel, Orden, Tipo, Documento, Renglon,")
			loConsulta.AppendLine("                        Estatus, Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori, Ren_Ori,")
			loConsulta.AppendLine("                        Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                        Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, Aud_Detalle)")
			loConsulta.AppendLine("SELECT      (#tmpTraza.Nivel + 1)               AS Nivel,")
			loConsulta.AppendLine("            4                                   AS Orden,")
			loConsulta.AppendLine("            'Recepciones'                       AS Tipo,")
			loConsulta.AppendLine("            Renglones_Recepciones.Documento     AS Documento,")
			loConsulta.AppendLine("            Renglones_Recepciones.Renglon       AS Renglon,")
			loConsulta.AppendLine("            Recepciones.Status                  AS Estatus,")
			loConsulta.AppendLine("            Recepciones.Fec_Ini                 AS Fec_Ini,")
			loConsulta.AppendLine("            Recepciones.Fec_Fin                 AS Fec_Fin,")
			loConsulta.AppendLine("            Renglones_Recepciones.Tip_Ori       AS Tip_Ori,")
			loConsulta.AppendLine("            Renglones_Recepciones.Doc_Ori       AS Doc_Ori,")
			loConsulta.AppendLine("            Renglones_Recepciones.Ren_Ori       AS Ren_Ori,")
			loConsulta.AppendLine("            Auditorias.Cod_Usu                  AS Aud_Cod_Usu,")
			loConsulta.AppendLine("            Auditorias.Registro                 AS Aud_Registro,")
			loConsulta.AppendLine("            Auditorias.Tipo                     AS Aud_Tipo,")
			loConsulta.AppendLine("            Auditorias.Tabla                    AS Aud_Tabla,")
			loConsulta.AppendLine("            Auditorias.Cod_Obj                  AS Aud_Cod_Obj,")
			loConsulta.AppendLine("            Auditorias.Notas                    AS Aud_Notas,")
			loConsulta.AppendLine("            Auditorias.Opcion                   AS Aud_Opcion,")
			loConsulta.AppendLine("            Auditorias.Accion                   AS Aud_Accion,")
			loConsulta.AppendLine("            Auditorias.Equipo                   AS Aud_Equipo,")
			loConsulta.AppendLine("            Auditorias.Detalle                  AS Aud_Detalle")
			loConsulta.AppendLine("FROM        #tmpTraza")
			loConsulta.AppendLine("    JOIN    Renglones_Recepciones ")
			loConsulta.AppendLine("        ON  Renglones_Recepciones.Tip_Ori = #tmpTraza.Tipo")
			loConsulta.AppendLine("        AND Renglones_Recepciones.Doc_Ori = #tmpTraza.Documento")
			loConsulta.AppendLine("        AND Renglones_Recepciones.Ren_Ori = #tmpTraza.Renglon")
			loConsulta.AppendLine("    JOIN    Recepciones")
			loConsulta.AppendLine("        ON  Recepciones.Documento = Renglones_Recepciones.Documento")
			loConsulta.AppendLine("    LEFT JOIN Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Recepciones.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Recepciones' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('NotasRecepcion', 'Sin opción');")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- Facturas de Compra")
			loConsulta.AppendLine("INSERT INTO #tmpTraza(  Nivel, Orden, Tipo, Documento, Renglon,")
			loConsulta.AppendLine("                        Estatus, Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori, Ren_Ori,")
			loConsulta.AppendLine("                        Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                        Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, Aud_Detalle)")
			loConsulta.AppendLine("SELECT      (#tmpTraza.Nivel + 1)               AS Nivel,")
			loConsulta.AppendLine("            5                                   AS Orden,")
			loConsulta.AppendLine("            'Compras'                           AS Tipo,")
			loConsulta.AppendLine("            Renglones_Compras.Documento         AS Documento,")
			loConsulta.AppendLine("            Renglones_Compras.Renglon           AS Renglon,")
			loConsulta.AppendLine("            Compras.Status                      AS Estatus,")
			loConsulta.AppendLine("            Compras.Fec_Ini                     AS Fec_Ini,")
			loConsulta.AppendLine("            Compras.Fec_Fin                     AS Fec_Fin,")
			loConsulta.AppendLine("            Renglones_Compras.Tip_Ori           AS Tip_Ori,")
			loConsulta.AppendLine("            Renglones_Compras.Doc_Ori           AS Doc_Ori,")
			loConsulta.AppendLine("            Renglones_Compras.Ren_Ori           AS Ren_Ori,")
			loConsulta.AppendLine("            Auditorias.Cod_Usu                  AS Aud_Cod_Usu,")
			loConsulta.AppendLine("            Auditorias.Registro                 AS Aud_Registro,")
			loConsulta.AppendLine("            Auditorias.Tipo                     AS Aud_Tipo,")
			loConsulta.AppendLine("            Auditorias.Tabla                    AS Aud_Tabla,")
			loConsulta.AppendLine("            Auditorias.Cod_Obj                  AS Aud_Cod_Obj,")
			loConsulta.AppendLine("            Auditorias.Notas                    AS Aud_Notas,")
			loConsulta.AppendLine("            Auditorias.Opcion                   AS Aud_Opcion,")
			loConsulta.AppendLine("            Auditorias.Accion                   AS Aud_Accion,")
			loConsulta.AppendLine("            Auditorias.Equipo                   AS Aud_Equipo,")
			loConsulta.AppendLine("            Auditorias.Detalle                  AS Aud_Detalle")
			loConsulta.AppendLine("FROM        #tmpTraza")
			loConsulta.AppendLine("    JOIN    Renglones_Compras ")
			loConsulta.AppendLine("        ON  Renglones_Compras.Tip_Ori = #tmpTraza.Tipo")
			loConsulta.AppendLine("        AND Renglones_Compras.Doc_Ori = #tmpTraza.Documento")
			loConsulta.AppendLine("        AND Renglones_Compras.Ren_Ori = #tmpTraza.Renglon")
			loConsulta.AppendLine("    JOIN    Compras")
			loConsulta.AppendLine("        ON  Compras.Documento = Renglones_Compras.Documento")
			loConsulta.AppendLine("    LEFT JOIN Auditorias ")
			loConsulta.AppendLine("        ON  Auditorias.Documento = Compras.Documento")
			loConsulta.AppendLine("		AND	Auditorias.Tabla	=	'Compras' ")
			loConsulta.AppendLine("		AND	Auditorias.Tipo		=	'Datos' ")
			loConsulta.AppendLine("		AND	Auditorias.Opcion	IN ('FacturasCompra', 'Sin opción');")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("-- ******************************************************************************")
			loConsulta.AppendLine("-- SELECT FINAL")
			loConsulta.AppendLine("-- ******************************************************************************")
			loConsulta.AppendLine("SELECT  Datos.Nivel                                             AS Nivel, ")
			loConsulta.AppendLine("        Datos.Orden                                             AS Orden, ")
			loConsulta.AppendLine("        Datos.Tipo                                              AS Tipo, ")
			loConsulta.AppendLine("        Datos.Documento                                         AS Documento, ")
			loConsulta.AppendLine("        Datos.Estatus                                           AS Estatus, ")
			loConsulta.AppendLine("        Datos.Fec_Ini                                           AS Fec_Ini, ")
			loConsulta.AppendLine("        Datos.Fec_Fin                                           AS Fec_Fin, ")
			loConsulta.AppendLine("        Datos.Dias                                              AS Dias, ")
			loConsulta.AppendLine("        Datos.Tip_Ori                                           AS Tip_Ori, ")
			loConsulta.AppendLine("        Datos.Doc_Ori                                           AS Doc_Ori,")
			loConsulta.AppendLine("        Datos.Aud_Cod_Usu                                       AS Aud_Cod_Usu, ")
			loConsulta.AppendLine("        Datos.Aud_Registro                                      AS Aud_Registro, ")
			loConsulta.AppendLine("        Datos.Aud_Tipo                                          AS Aud_Tipo, ")
			loConsulta.AppendLine("        Datos.Aud_Tabla                                         AS Aud_Tabla, ")
			loConsulta.AppendLine("        Datos.Aud_Cod_Obj                                       AS Aud_Cod_Obj, ")
			loConsulta.AppendLine("        Datos.Aud_Notas                                         AS Aud_Notas,")
			loConsulta.AppendLine("        Datos.Aud_Opcion                                        AS Aud_Opcion, ")
			loConsulta.AppendLine("        Datos.Aud_Accion                                        AS Aud_Accion, ")
			loConsulta.AppendLine("        Datos.Aud_Equipo                                        AS Aud_Equipo, ")
			loConsulta.AppendLine("        Aud_Encabezado.C.value('@nombre[1]', 'VARCHAR(MAX)')    AS Campo_Nombre,")
			loConsulta.AppendLine("        Aud_Encabezado.C.value('./antes[1]', 'VARCHAR(MAX)')    AS Campo_Antes,")
			loConsulta.AppendLine("        Aud_Encabezado.C.value('./despues[1]', 'VARCHAR(MAX)')  AS Campo_Despues,")
			loConsulta.AppendLine("        Aud_Detalle2")
			loConsulta.AppendLine("FROM  ( SELECT      Nivel, Orden, Tipo, Documento, Estatus, Fec_Ini, Fec_Fin, ")
			loConsulta.AppendLine("                    DATEDIFF(DAY, Fec_Ini, Fec_Fin) Dias, Tip_Ori, Doc_Ori,")
			loConsulta.AppendLine("                    Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj, ")
			loConsulta.AppendLine("                    Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, ")
			loConsulta.AppendLine("                    Aud_Detalle Aud_Detalle2,")
			loConsulta.AppendLine("                    CAST(Aud_Detalle AS XML) Aud_Detalle")
			loConsulta.AppendLine("        FROM        #tmpTraza")
			loConsulta.AppendLine("        GROUP BY    Nivel, Orden, Tipo, Documento, Estatus, ")
			loConsulta.AppendLine("                    Fec_Ini, Fec_Fin, Tip_Ori, Doc_Ori,")
			loConsulta.AppendLine("                    Aud_Cod_Usu, Aud_Registro, Aud_Tipo, Aud_Tabla, Aud_Cod_Obj,")
			loConsulta.AppendLine("                    Aud_Notas, Aud_Opcion, Aud_Accion, Aud_Equipo, ")
			loConsulta.AppendLine("                    Aud_Detalle                    ")
			loConsulta.AppendLine("    ) Datos")
			loConsulta.AppendLine("    OUTER APPLY Datos.Aud_Detalle.nodes('detalle/campos/campo') AS Aud_Encabezado(C)")
			loConsulta.AppendLine("ORDER BY    Nivel, Orden, Documento, Aud_Registro;")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
		    
            Dim loServicios As New cusDatos.goDatos()
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

           '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAuditoriasTraza_Requisiciones", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
            Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvfAuditoriasTraza_Requisiciones.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 11/12/14: Codigo inicial, a partir de fAuditorias_Requisiciones.                     '
'-------------------------------------------------------------------------------------------'
