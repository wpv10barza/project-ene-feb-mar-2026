import os
import sys

# Añadir el path actual para que reconozca los módulos
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from modules.ingestor import DocumentIngestor
from modules.retriever import DocumentRetriever

BASE_DIR = r"D:\118102025\Documents\trabajos\RPA\STUDIOLM"
DATA_DIR = os.path.join(BASE_DIR, "mis_documentos")
DB_DIR = os.path.join(BASE_DIR, "db_gemma_rag")

def main():
    print("\n=== RPA DOCUMENT MANAGER (RTX 3050 READY) ===")
    print("1. Indexar documentos (PDFs -> VectorStore)")
    print("2. Consultar con Filtro (Actas/Transcripciones)")
    
    op = input("\nElige una opcion: ")

    if op == "1":
        ingestor = DocumentIngestor(DATA_DIR, DB_DIR)
        print(ingestor.run())
    elif op == "2":
        cat = input("¿En que categoria buscar? (actas/transcripciones/guias): ").lower()
        query = input(f"Pregunta para {cat}: ")
        retriever = DocumentRetriever(DB_DIR)
        contexto = retriever.get_context(query, cat)
        print("\n--- CONTEXTO ENCONTRADO ---")
        print(contexto if contexto else "No se encontró información relevante.")

if __name__ == "__main__":
    main()
