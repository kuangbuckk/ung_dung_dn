// server.js
const express = require('express');
const { GoogleGenAI } = require('@google/genai');
const cors = require('cors');
require('dotenv').config();

// Lấy khóa API
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "YOUR_FALLBACK_API_KEY"; 
const ai = new GoogleGenAI(GEMINI_API_KEY);

const app = express();
const port = 3001; 

app.use(cors()); 
app.use(express.json());

// --- Cấu hình Tool Google Search (Cho Giải thích Thuật ngữ) ---
const searchTool = {
    tools: [{ googleSearch: {} }],
};
// -----------------------------------------------------------

// =======================================================
// A. ENDPOINT CŨ: Xử lý yêu cầu tìm kiếm NLP
// =======================================================
app.post('/api/process-search', async (req, res) => {
    const { naturalQuery } = req.body;

    if (!naturalQuery) {
        return res.status(400).json({ error: "Thiếu 'naturalQuery' trong body." });
    }

    const prompt = `
        Bạn là một trình phân tích ngôn ngữ tự nhiên. 
        Hãy chuyển đổi yêu cầu tìm kiếm bằng tiếng Việt sau thành một chuỗi JSON thuần túy (RAW JSON) mà tôi có thể dùng để tìm kiếm trong tài liệu Word. 
        Trích xuất các từ khóa/cụm từ quan trọng nhất. Nếu người dùng muốn tìm kiếm chính xác, hãy đặt matchWholeWord là true.
        
        Yêu cầu tìm kiếm: "${naturalQuery}"

        ĐỊNH DẠNG ĐẦU RA PHẢI LÀ JSON NGUYÊN BẢN:
        {
          "keywords": ["từ khóa 1", "cụm từ 2"],
          "options": {
            "matchWholeWord": true/false 
          }
        }
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash', 
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            config: {
                temperature: 0.1, 
                responseMimeType: "application/json" 
            }
        });

        const jsonText = response.text.trim().replace(/```json|```/g, '');
        const searchPlan = JSON.parse(jsonText);

        res.json(searchPlan);
    } catch (error) {
        console.error("Lỗi khi gọi Gemini (NLP Search):", error);
        res.status(500).json({ error: "Lỗi nội bộ, lỗi phân tích JSON, hoặc API Key không hợp lệ." });
    }
});


// =======================================================
// B. ENDPOINT MỚI 1: Tóm tắt Tài liệu (Document Summarization)
// =======================================================
app.post('/api/summarize', async (req, res) => {
    const { documentText, summaryRequest } = req.body;

    if (!documentText) {
        return res.status(400).json({ error: "Thiếu 'documentText' để tóm tắt." });
    }

    const prompt = `
        Tóm tắt đoạn văn bản sau dựa trên yêu cầu: "${summaryRequest || 'Tóm tắt các ý chính'}"
        
        Văn bản:
        ---
        ${documentText}
        ---
        
        Hãy trả lời bằng tiếng Việt.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash', 
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            config: {
                temperature: 0.3,
            }
        });

        res.json({ result: response.text });
    } catch (error) {
        console.error("Lỗi khi gọi Gemini (Summarize):", error);
        res.status(500).json({ error: "Lỗi nội bộ khi gọi API Gemini." });
    }
});


// =======================================================
// C. ENDPOINT MỚI 2: Hỏi & Đáp Theo Ngữ Cảnh (Contextual Q&A)
// =======================================================
app.post('/api/qna', async (req, res) => {
    const { documentText, userQuestion } = req.body;

    if (!documentText || !userQuestion) {
        return res.status(400).json({ error: "Thiếu 'documentText' hoặc 'userQuestion' cho Q&A." });
    }

    const prompt = `
        Bạn là một trợ lý thông minh. Hãy trả lời câu hỏi sau chỉ dựa trên NGỮ CẢNH được cung cấp dưới đây. 
        Nếu thông tin KHÔNG CÓ trong ngữ cảnh, hãy nói "Thông tin không có trong tài liệu này".
        
        Câu hỏi: "${userQuestion}"
        
        NGỮ CẢNH:
        ---
        ${documentText}
        ---
        
        Hãy trả lời bằng tiếng Việt.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash', 
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            config: {
                temperature: 0.1, 
            }
        });

        res.json({ result: response.text });
    } catch (error) {
        console.error("Lỗi khi gọi Gemini (QNA):", error);
        res.status(500).json({ error: "Lỗi nội bộ khi gọi API Gemini." });
    }
});


// =======================================================
// D. ENDPOINT MỚI 3: Giải thích Thuật ngữ Kèm Theo Nghiên Cứu
// =======================================================
app.post('/api/explain', async (req, res) => {
    const { term } = req.body;

    if (!term) {
        return res.status(400).json({ error: "Thiếu 'term' cần giải thích." });
    }

    const prompt = `
        Giải thích thuật ngữ "${term}" một cách rõ ràng và ngắn gọn. Sau đó, tìm kiếm trên Google (nếu cần) và cung cấp thêm 1-2 thông tin liên quan, cập nhật hoặc ví dụ thực tế.
        
        Hãy trả lời bằng tiếng Việt.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash', 
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            config: {
                temperature: 0.3,
                ...searchTool // Kích hoạt công cụ tìm kiếm Google
            }
        });

        res.json({ result: response.text });
    } catch (error) {
        console.error("Lỗi khi gọi Gemini (Explain):", error);
        res.status(500).json({ error: "Lỗi nội bộ khi gọi API Gemini." });
    }
});


app.listen(port, () => {
    console.log(`Backend server đang chạy tại http://localhost:${port}`);
});