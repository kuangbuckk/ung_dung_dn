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
// A. ENDPOINT 1: Hỏi & Đáp Theo Ngữ Cảnh (Contextual Q&A)
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
// B. ENDPOINT 2: Giải thích Thuật ngữ Kèm Theo Nghiên Cứu
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