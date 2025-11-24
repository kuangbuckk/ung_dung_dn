// ===============================
// CONFIG
// ===============================
const BACKEND_URL = "http://localhost:3001/api";
let fileId = null;
let fileDisplayName = "Document_Content.txt";

// ===============================
// UI HELPERS
// ===============================
const BUTTON_LABELS = {
    "index-button": "Upload Tài liệu Lên File Store",
    "qna-button": "Trả lời Câu hỏi (Toàn Tài liệu)",
    "explain-button": "Giải thích Thuật ngữ (Kèm Nghiên Cứu)"
};

function setProcessing(isProcessing, button, statusDiv, processName) {
    const spinner = document.getElementById("loading-spinner");

    Object.entries(BUTTON_LABELS).forEach(([id, text]) => {
        const btn = document.getElementById(id);
        if (!btn) return;

        btn.disabled = isProcessing;
        btn.textContent =
            isProcessing && btn === button
                ? `Đang xử lý ${processName}...`
                : text;
    });

    if (isProcessing) {
        spinner.classList.remove("d-none");
        statusDiv.textContent = `Đang thực hiện ${processName} bằng Gemini...`;
    } else {
        spinner.classList.add("d-none");
    }
}

function displayResult(resultsDiv, title, content) {
    resultsDiv.innerHTML = `
        <h4 style="padding-bottom:6px;border-bottom:1px solid #ddd">${title}</h4>
        <p style="white-space:pre-wrap">${content}</p>
    `;
}

// ===============================
// WORD API HELPERS
// ===============================
async function getFullDocumentText() {
    let fullText = "";
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        fullText = body.text.trim();
    });
    return fullText;
}

async function getSelectedText(context) {
    const r = context.document.getSelection();
    r.load("text");
    await context.sync();
    return r.text.trim();
}

// ===============================
// BACKEND CALL WRAPPERS
// ===============================
async function apiPost(path, data) {
    const res = await fetch(`${BACKEND_URL}/${path}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });

    const json = await res.json().catch(() => ({
        error: `HTTP ${res.status}`,
    }));

    if (!res.ok || json.error) {
        throw new Error(json.error || `HTTP ${res.status}`);
    }
    return json;
}

// UPLOAD
async function runUploadLogic(documentText, displayName) {
    const res = await apiPost("upload-file", {
        documentText,
        displayName,
    });

    fileId = res.fileId;
    return res.message;
}

// Global Q&A
async function runGlobalQNALogic(question, resultsDiv) {
    const data = await apiPost("global-qna", { userQuestion: question });

    let text = data.result;
    if (data.citation) text += `\n\n--- Trích dẫn ---\n${data.citation}`;

    displayResult(resultsDiv, `Kết quả cho câu hỏi: "${question}"`, text);
}

// Explain
async function runExplainLogic(term, resultsDiv) {
    const data = await apiPost("explain", { term });
    displayResult(resultsDiv, `Giải thích: "${term}"`, data.result);
}

// ===============================
// EVENT WRAPPERS
// ===============================
async function runUpload(statusDiv, button) {
    const process = "Upload File";
    setProcessing(true, button, statusDiv, process);

    try {
        const text = await getFullDocumentText();
        if (!text) throw new Error("Tài liệu trống.");

        const msg = await runUploadLogic(text, fileDisplayName);
        statusDiv.textContent = `${msg}. ID: ${fileId}`;

    } catch (e) {
        statusDiv.textContent = `Lỗi Upload: ${e.message}`;
    } finally {
        setProcessing(false, button, statusDiv, process);
    }
}

async function runGlobalQNA(statusDiv, resultsDiv, button, input) {
    const process = "Global Q&A";
    setProcessing(true, button, statusDiv, process);

    try {
        const q = input.value.trim();
        if (!q) throw new Error("Vui lòng nhập câu hỏi.");
        if (!fileId) throw new Error("Chưa upload tài liệu.");

        await runGlobalQNALogic(q, resultsDiv);
        statusDiv.textContent = "Hoàn tất Global Q&A.";
    } catch (e) {
        statusDiv.textContent = `Lỗi: ${e.message}`;
    } finally {
        setProcessing(false, button, statusDiv, process);
    }
}

async function runExplain(statusDiv, resultsDiv, button) {
    const process = "Giải thích";
    setProcessing(true, button, statusDiv, process);

    await Word.run(async (context) => {
        try {
            const term = await getSelectedText(context);
            if (!term) throw new Error("Vui lòng chọn thuật ngữ.");

            await runExplainLogic(term, resultsDiv);
            statusDiv.textContent = "Hoàn tất.";
        } catch (e) {
            statusDiv.textContent = `Lỗi: ${e.message}`;
        }
    }).finally(() => {
        setProcessing(false, button, statusDiv, process);
    });
}

// ===============================
// INIT
// ===============================
Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) return;

    const statusDiv = document.getElementById("status");
    const resultsDiv = document.getElementById("results-content");

    const btnUpload = document.getElementById("index-button");
    const btnQna = document.getElementById("qna-button");
    const btnExplain = document.getElementById("explain-button");
    const inputQna = document.getElementById("qna-input");

    // Auto-upload khi mở
    statusDiv.textContent = "Khởi tạo... Đang upload tài liệu...";
    setTimeout(() => runUpload(statusDiv, btnUpload), 500);

    btnUpload.onclick = () => runUpload(statusDiv, btnUpload);
    btnQna.onclick = () => runGlobalQNA(statusDiv, resultsDiv, btnQna, inputQna);
    btnExplain.onclick = () => runExplain(statusDiv, resultsDiv, btnExplain);
});
