import os
import time
import json
import logging
import numpy as np
from flask import Flask, request, jsonify
from flask_cors import CORS
from dotenv import load_dotenv
from tqdm import tqdm

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
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")
if not GOOGLE_API_KEY:
    raise Exception("Missing GOOGLE_API_KEY in .env")

# ----------------------------
# Embedding Model
# ----------------------------
model = SentenceTransformer("all-MiniLM-L6-v2")

class CustomEmbeddings(Embeddings):
    def __init__(self, model):
        self.model = model

    def embed_documents(self, texts):
        return self.model.encode(
            texts, batch_size=64, show_progress_bar=True
        ).tolist()

    def embed_query(self, text):
        vec = self.model.encode([text], show_progress_bar=False)
        return vec[0].tolist()

embeddings = CustomEmbeddings(model)

# ----------------------------
# Vector DB & metadata
# ----------------------------
vector_db = None
FAISS_PATH = "faiss_index"
METADATA_PATH = "metadata.json"
FILE_STORAGE_PATH = "uploaded_document.txt"

def load_vector_db():
    global vector_db
    if os.path.exists(FAISS_PATH) and os.path.exists(METADATA_PATH):
        try:
            vector_db = FAISS.load_local(FAISS_PATH, embeddings)
            logging.info(f"Loaded FAISS index from disk")
        except Exception as e:
            logging.error(f"Failed to load FAISS: {e}")
            vector_db = None
    else:
        vector_db = None

def improved_splitter():
    return RecursiveCharacterTextSplitter(
        separators=["\n\n", ". ", "\n"],
        chunk_size=1000,
        chunk_overlap=200,
        length_function=len,
        keep_separator=True
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
- Nếu không tìm thấy thông tin liên quan, hãy trả lời rằng bạn không biết.

Trả lời bằng tiếng Việt.
"""
prompt_template = PromptTemplate.from_template(PROMPT_TEMPLATE)

# Gemini LLM
llm = ChatGoogleGenerativeAI(
    model="gemini-2.5-flash",
    temperature=0.1
)

# ----------------------------
# UPLOAD + CHUNK + INDEX
# ----------------------------
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
        new_chunks = splitter.split_documents(documents)

        # Load existing vector DB
        load_vector_db()

        # Incremental FAISS với batch embedding
        batch_size = 64
        total_batches = (len(new_chunks) + batch_size - 1) // batch_size
        logging.info(f"Total chunks: {len(new_chunks)}, batch_size: {batch_size}, total_batches: {total_batches}")

        for i in tqdm(range(0, len(new_chunks), batch_size),
                      total=total_batches,
                      desc="Batches",
                      ncols=100,
                      unit="batch"):
            batch = new_chunks[i:i+batch_size]
            if vector_db:
                vector_db.add_documents(batch)
            else:
                vector_db = FAISS.from_documents(batch, embeddings)
            logging.info(f"Processed batch {i//batch_size + 1}/{total_batches}")

        # Save
        vector_db.save_local(FAISS_PATH)
        with open(METADATA_PATH, "w", encoding="utf-8") as f:
            json.dump([d.page_content for d in new_chunks], f, ensure_ascii=False)

        total_time = round(time.time() - t0, 4)
        logging.info(f"Indexed {len(new_chunks)} chunks in {total_time}s")

        return jsonify({
            "message": "Upload & indexing thành công!",
            "chunks": len(new_chunks),
            "time": total_time,
            "fileId": "local-faiss"
        })

    except Exception as e:
        logging.error(f"Upload failed: {e}")
        return jsonify({"error": str(e)}), 500

# ----------------------------
# GLOBAL Q&A — semantic + hybrid search
# ----------------------------
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
        retriever = vector_db.as_retriever(search_type="similarity", search_kwargs={"k": 15})
        top_docs = retriever.invoke(question)

        query_vec = np.array(embeddings.embed_query(question))
        results = []
        for i, d in enumerate(top_docs):
            doc_vec = np.array(embeddings.embed_query(d.page_content))
            cosine_sim = float(np.dot(query_vec, doc_vec) / (np.linalg.norm(query_vec) * np.linalg.norm(doc_vec) + 1e-8))
            token_overlap = len(set(question.split()) & set(d.page_content.split())) / max(1, len(set(question.split())))
            if cosine_sim >= 0.35 and token_overlap >= 0.1:
                results.append({"chunkIndex": i, "score": cosine_sim, "text": d.page_content})

        results = sorted(results, key=lambda x: -x["score"])[:10]  # lấy top 10
        context = "\n\n---\n\n".join([r["text"] for r in results]) or "Không tìm thấy thông tin liên quan."

        t1 = time.time()
        prompt = prompt_template.format(context=context, question=question)
        response = llm.invoke(prompt)
        llm_time = round(time.time() - t1, 4)
        total_time = round(time.time() - full_start, 4)

        logging.info(f"Q: {question} | Top Chunks: {len(results)} | LLM: {llm_time}s | Total: {total_time}s")

        return jsonify({
            "result": response.content,
            "citations": results[:1],  # highlight top 1
            "timing": {"llm": llm_time, "total": total_time}
        })

    except Exception as e:
        logging.error(f"Global QnA failed: {e}")
        return jsonify({"error": str(e)}), 500

# ----------------------------
# MAIN
# ----------------------------
if __name__ == "__main__":
    load_vector_db()  # Load FAISS + metadata nếu có
    app.run(host="0.0.0.0", port=3001, debug=True)
