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

        df_map = validar_mapeo(df_excel, df_preg, strict=False)

        codigos_excel = set(df_map["PREGUNTA_CODIGO"].astype(str).str.strip().tolist())
        codigos_bd = set(df_preg["PREGUNTA_CODIGO"].astype(str).str.strip().tolist())

        print(f"EvaluacionId={args.evaluacion_id}")
        print(f"- Preguntas en BD: {total_preg}")
        print(f"- Códigos en Excel: {len(codigos_excel)}")
        print(f"- Códigos con match: {len(codigos_excel & codigos_bd)}")
        print(f"- Códigos extra en Excel: {len(codigos_excel - codigos_bd)}")

        cargables = df_map[~df_map["ID_PREGUNTA"].isna()]
        nocargables = df_map[df_map["ID_PREGUNTA"].isna()]
        print(f"- Filas cargables: {len(cargables)}")
        print(f"- Filas NO cargables: {len(nocargables)}")

        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)
        print(f"AplicacionId: {aplicacion_id}")

        n_resp = contar_respuestas(conn, aplicacion_id)
        respondidas = preguntas_respondidas(conn, aplicacion_id)
        print(f"- Respuestas existentes: {n_resp}")
        print(f"- Preguntas subidas: {len(respondidas)}")
        print(f"- Preguntas no subidas: {max(0, total_preg - len(respondidas))}")
    return 0

def cmd_load_auto(args) -> int:
    df_excel = leer_excel([args.excel], sheet_name=args.sheet)
    with get_connection() as conn:
        df_preg = cargar_preguntas(conn, args.evaluacion_id)
        aplicacion_id = obtener_aplicacion_id(conn, args.evaluacion_id, args.informante_id)

        # CARGA: estricto
        df_map = validar_mapeo(df_excel, df_preg, strict=True)

        n_exist = contar_respuestas(conn, aplicacion_id)
        if n_exist == 0:
            n = insertar_respuestas_solo_faltantes(conn, aplicacion_id, df_map, args.usuario_registro_id)
            print(f"[OK] Tabla vacía -> INSERT. Nuevos registros: {n}")
        else:
            stats = merge_concatenar_respuestas(conn, aplicacion_id, df_map, args.usuario_registro_id)
            print(f"[OK] Tabla con datos -> MERGE. INSERT={stats.get('INSERT',0)} | UPDATE={stats.get('UPDATE',0)}")
    return 0

def build_parser():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)

    def add_common(sp):
        sp.add_argument("--evaluacion-id", type=int, default=DEFAULT_EVALUACION_ID)
        sp.add_argument("--informante-id", type=int, default=DEFAULT_INFORMANTE_ID)
        sp.add_argument("--usuario-registro-id", type=int, default=DEFAULT_USUARIO_REGISTRO_ID)
        sp.add_argument("--excel", required=True)
        sp.add_argument("--sheet", default="RESULTADO")

    sp = sub.add_parser("preview")
    add_common(sp)
    sp.set_defaults(func=cmd_preview)

    sp = sub.add_parser("load-auto")
    add_common(sp)
    sp.set_defaults(func=cmd_load_auto)

    return p

def main():
    args = build_parser().parse_args()
    return args.func(args)

if __name__ == "__main__":
    raise SystemExit(main())

