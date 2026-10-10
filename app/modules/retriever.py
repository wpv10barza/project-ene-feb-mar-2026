from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

class DocumentRetriever:
    def __init__(self, db_path):
        self.embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
        self.db = Chroma(persist_directory=db_path, embedding_function=self.embeddings)

    def get_context(self, query, category):
        # Filtro inteligente por metadatos
        docs = self.db.similarity_search(query, k=3, filter={"tipo_doc": category})
        return "\n---\n".join([d.page_content for d in docs])
