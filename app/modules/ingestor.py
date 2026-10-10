import os
from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

class DocumentIngestor:
    def __init__(self, data_path, db_path):
        self.data_path = data_path
        self.db_path = db_path
        self.embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")

    def run(self):
        docs_metadata = []
        for cat in ["actas", "transcripciones", "guias"]:
            folder = os.path.join(self.data_path, cat)
            if not os.path.exists(folder): continue
            
            for file in os.listdir(folder):
                if file.endswith(".pdf"):
                    print(f"Cargando {file}...")
                    loader = PyPDFLoader(os.path.join(folder, file))
                    pages = loader.load()
                    for p in pages:
                        p.metadata.update({"tipo_doc": cat, "fuente": file})
                    docs_metadata.extend(pages)

        if not docs_metadata: return "No hay archivos PDF en las carpetas."

        splitter = RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=100)
        chunks = splitter.split_documents(docs_metadata)
        
        Chroma.from_documents(documents=chunks, embedding=self.embeddings, persist_directory=self.db_path)
        return f"Proceso terminado: {len(chunks)} fragmentos guardados en DB."
