import os
import time
import logging
import numpy as np
from flask import Flask, request, jsonify
from flask_cors import CORS
from dotenv import load_dotenv

# RAG Components
from langchain_community.vectorstores import FAISS
from langchain_community.document_loaders import TextLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import PromptTemplate
from langchain_core.embeddings import Embeddings
from sentence_transformers import SentenceTransformer

load_dotenv()

# ----------------------------
# INIT APP
# ----------------------------
app = Flask(__name__)
CORS(app)

logging.basicConfig(
    filename="rag.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
if not GOOGLE_API_KEY:
    raise Exception("Missing GOOGLE_API_KEY in .env")


# ----------------------------
# Embedding Model
# ----------------------------
base_model = SentenceTransformer("all-MiniLM-L6-v2")

class CustomSentenceTransformerEmbeddings(Embeddings):
    def __init__(self, model):
        self.model = model

    def embed_documents(self, texts):
        return self.model.encode(texts).tolist()

    def embed_query(self, text):
        return self.model.encode(text).tolist()

embeddings = CustomSentenceTransformerEmbeddings(base_model)

# ----------------------------
# Vector DB
# ----------------------------
vector_db = None
FILE_STORAGE_PATH = "uploaded_document.txt"


# ----------------------------
# Improve Chunking
# ----------------------------
def improved_splitter():
    return RecursiveCharacterTextSplitter(
        separators=["\n\n", "\n", ". ", " "],
        chunk_size=500,
        chunk_overlap=100,
        length_function=len,
        keep_separator=False
    )


# ----------------------------
# RAG Prompt
# ----------------------------
PROMPT_TEMPLATE = """
Bạn là trợ lý AI RAG. Chỉ sử dụng thông tin từ tài liệu để trả lời.

--- NGỮ CẢNH ---
{context}

--- Câu hỏi ---
{question}

YÊU CẦU:
- Trả lời rõ ràng, đúng sự thật.
- Không bịa đặt nếu thiếu thông tin.
- Nếu câu trả lời đến từ nhiều phần của tài liệu, hãy tổng hợp lại.

Trả lời bằng tiếng Việt.
"""
prompt_template = PromptTemplate.from_template(PROMPT_TEMPLATE)

# Gemini LLM
llm = ChatGoogleGenerativeAI(
    model="gemini-2.5-flash",
    temperature=0.2
)

# ================================
# 🔥 A. UPLOAD + CHUNK + INDEX
# ================================
@app.route("/api/upload-file", methods=["POST"])
def upload_file():
    global vector_db

    t0 = time.time()

    data = request.json
    document_text = data.get("documentText", "")
    display_name = data.get("displayName", "document.txt")

    if not document_text.strip():
        return jsonify({"error": "Document content is empty"}), 400

    # Save document
    with open(FILE_STORAGE_PATH, "w", encoding="utf-8") as f:
        f.write(document_text)

    try:
        loader = TextLoader(FILE_STORAGE_PATH, encoding="utf-8")
        documents = loader.load()

        splitter = improved_splitter()
        final_docs = splitter.split_documents(documents)

        # Rebuild FAISS
        vector_db = FAISS.from_documents(final_docs, embeddings)
        vector_db.save_local("faiss_index")

        total_time = round(time.time() - t0, 4)

        logging.info(f"INDEXED {len(final_docs)} chunks in {total_time}s")

        return jsonify({
            "message": "Upload & indexing thành công!",
            "chunks": len(final_docs),
            "time": total_time,
            "fileId": "local-faiss"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# =====================================================
# 🔥 B. GLOBAL Q&A — semantic filtering + citations
# =====================================================
@app.route("/api/global-qna", methods=["POST"])
def global_qna():
    global vector_db

    if vector_db is None:
        return jsonify({"error": "Vector DB chưa được tạo. Hãy upload tài liệu trước."}), 400

    data = request.json
    question = data.get("userQuestion", "")

    if not question:
        return jsonify({"error": "Missing question"}), 400

    full_start = time.time()

    try:
        # ---- Search ----
        t0 = time.time()
        retriever = vector_db.as_retriever(search_type="similarity", search_kwargs={"k": 10})
        docs = retriever.invoke(question)
        search_time = round(time.time() - t0, 4)

        # ---- Semantic Filtering ----
        query_vec = embeddings.embed_query(question)
        results = []

        for i, d in enumerate(docs):
            doc_vec = embeddings.embed_query(d.page_content)
            score = np.dot(query_vec, doc_vec) / (np.linalg.norm(query_vec) * np.linalg.norm(doc_vec))

            if score >= 0.35:     # threshold — tuneable
                results.append({
                    "chunkIndex": i,
                    "score": float(score),
                    "text": d.page_content[:800],
                })

        # Sort by semantic relevance
        results = sorted(results, key=lambda x: -x["score"])[:5]

        # Build context for LLM
        context = "\n\n---\n\n".join([r["text"] for r in results]) or "Không tìm thấy thông tin liên quan."

        # ---- LLM ----
        t1 = time.time()
        prompt = prompt_template.format(context=context, question=question)
        response = llm.invoke(prompt)
        llm_time = round(time.time() - t1, 4)

        # Total time
        total = round(time.time() - full_start, 4)

        # LOGGING
        logging.info(f"""
======== RAG QUERY ========
Question: {question}
Context Chunks: {len(results)}
Search Time: {search_time}s
LLM Time: {llm_time}s
Total: {total}s
Top Chunks: {results}
===========================
""")

        return jsonify({
            "result": response.content,
            "citations": results,
            "timing": {
                "search": search_time,
                "llm": llm_time,
                "total": total
            }
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===========================
# 🔥 C. Explain — plain LLM
# ===========================
@app.route("/api/explain", methods=["POST"])
def explain_term():
    data = request.json
    term = data.get("term", "")

    if not term:
        return jsonify({"error": "Missing term"}), 400

    try:
        prompt = f"""
Giải thích thuật ngữ sau bằng tiếng Việt, kèm ví dụ:

Thuật ngữ: {term}
"""
        response = llm.invoke(prompt)

        return jsonify({"result": response.content})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ----------------------------
# MAIN
# ----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3001, debug=True)
