import argparse
from config import DEFAULT_EVALUACION_ID, DEFAULT_INFORMANTE_ID, DEFAULT_USUARIO_REGISTRO_ID
from src.db import get_connection
from src.excel_reader import leer_excel
from src.mapping import validar_mapeo
from src.repo import cargar_preguntas, obtener_aplicacion_id, contar_respuestas, preguntas_respondidas
from src.loaders import insertar_respuestas_solo_faltantes, merge_concatenar_respuestas

def cmd_preview(args) -> int:
    df_excel = leer_excel([args.excel], sheet_name=args.sheet)
    with get_connection() as conn:
        df_preg = cargar_preguntas(conn, args.evaluacion_id)
        total_preg = len(df_preg)

        df_map = validar_mapeo(df_excel, df_preg)
        codigos_excel = set(df_map["PREGUNTA_CODIGO"].astype(str).str.strip().tolist())
        codigos_bd = set(df_preg["PREGUNTA_CODIGO"].astype(str).str.strip().tolist())

        print(f"EvaluacionId={args.evaluacion_id}")
        print(f"- Preguntas en BD (VW_PREGUNTA_POR_EVALUACION): {total_preg}")
        print(f"- Códigos en Excel: {len(codigos_excel)}")
        print(f"- Códigos con match: {len(codigos_excel & codigos_bd)}")
        print(f"- Códigos que faltan en Excel (BD pero no Excel): {len(codigos_bd - codigos_excel)}")
        print(f"- Códigos extra en Excel (Excel pero no BD): {len(codigos_excel - codigos_bd)}")

        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)
        print(f"AplicacionId (EvaluacionId+InformanteId): {aplicacion_id}")

        n_resp = contar_respuestas(conn, aplicacion_id)
        respondidas = preguntas_respondidas(conn, aplicacion_id)

        print(f"- Respuestas existentes en dbo.Respuesta: {n_resp}")
        print(f"- Preguntas subidas (distinct PreguntaId): {len(respondidas)}")
        print(f"- Preguntas no subidas: {max(0, total_preg - len(respondidas))}")

    return 0

def cmd_load_insert(args) -> int:
    df_excel = leer_excel([args.excel], sheet_name=args.sheet)
    with get_connection() as conn:
        df_preg = cargar_preguntas(conn, args.evaluacion_id)
        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)
        df_map = validar_mapeo(df_excel, df_preg)

        n = insertar_respuestas_solo_faltantes(conn, aplicacion_id, df_map, args.usuario_registro_id)
        print(f"[OK] Insertadas (solo faltantes): {n}")
    return 0

def cmd_load_merge(args) -> int:
    df_excel = leer_excel([args.excel], sheet_name=args.sheet)
    with get_connection() as conn:
        df_preg = cargar_preguntas(conn, args.evaluacion_id)
        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)
        df_map = validar_mapeo(df_excel, df_preg)

        stats = merge_concatenar_respuestas(conn, aplicacion_id, df_map, args.usuario_registro_id)
        print(f"[OK] MERGE: INSERT={stats['INSERT']} | UPDATE={stats['UPDATE']}")
    return 0

def cmd_load_auto(args) -> int:
    df_excel = leer_excel([args.excel], sheet_name=args.sheet)
    with get_connection() as conn:
        df_preg = cargar_preguntas(conn, args.evaluacion_id)
        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)
        df_map = validar_mapeo(df_excel, df_preg)

        n_exist = contar_respuestas(conn, aplicacion_id)
        if n_exist == 0:
            n = insertar_respuestas_solo_faltantes(conn, aplicacion_id, df_map, args.usuario_registro_id)
            print(f"[OK] Tabla vacía -> INSERT. Nuevos registros añadidos: {n}")
        else:
            stats = merge_concatenar_respuestas(conn, aplicacion_id, df_map, args.usuario_registro_id)
            print(f"[OK] Tabla con datos -> MERGE. INSERT={stats['INSERT']} | UPDATE={stats['UPDATE']}")
    return 0

def build_parser():
    p = argparse.ArgumentParser(description="Cargador Evaluación (Excel -> SQL) con preview/insert/merge")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_common(sp):
        sp.add_argument("--evaluacion-id", type=int, default=DEFAULT_EVALUACION_ID)
        sp.add_argument("--informante-id", type=int, default=DEFAULT_INFORMANTE_ID)
        sp.add_argument("--usuario-registro-id", type=int, default=DEFAULT_USUARIO_REGISTRO_ID)
        sp.add_argument("--excel", required=True, help="Ruta al Excel")
        sp.add_argument("--sheet", default="RESULTADO")

    sp = sub.add_parser("preview", help="Preview (no escribe BD)")
    add_common(sp)
    sp.set_defaults(func=cmd_preview)

    sp = sub.add_parser("load-insert", help="Inserta solo faltantes")
    add_common(sp)
    sp.set_defaults(func=cmd_load_insert)

    sp = sub.add_parser("load-merge", help="Merge + concatenación")
    add_common(sp)
    sp.set_defaults(func=cmd_load_merge)

    sp = sub.add_parser("load-auto", help="Auto: si vacío inserta, si no merge")
    add_common(sp)
    sp.set_defaults(func=cmd_load_auto)

    return p

def main():
    args = build_parser().parse_args()
    return args.func(args)

if __name__ == "__main__":
    raise SystemExit(main())

