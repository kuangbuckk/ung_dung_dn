const express = require('express');
const { GoogleGenAI } = require('@google/genai');
const cors = require('cors');
require('dotenv').config();

// Lấy khóa API từ biến môi trường
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "YOUR_FALLBACK_API_KEY"; 
const ai = new GoogleGenAI(GEMINI_API_KEY);

const app = express();
const port = 3001;

app.use(cors()); // Cần thiết để frontend (localhost:3000) gọi API này
app.use(express.json());


app.post('/api/process-search', async (req, res) => {
    const { naturalQuery } = req.body;

    if (!naturalQuery) {
        return res.status(400).json({ error: "Thiếu 'naturalQuery'" });
    }

    const prompt = `
        Bạn là một trình phân tích ngôn ngữ tự nhiên. 
        Hãy chuyển đổi yêu cầu tìm kiếm bằng tiếng Việt sau thành một chuỗi JSON thuần túy (RAW JSON) mà tôi có thể dùng để tìm kiếm trong tài liệu Word.
        
        Bên cạnh việc trích xuất các từ khóa (keywords) để tìm kiếm, bạn phải **trích xuất một số cụm từ chính (relevant_phrases)** từ yêu cầu của người dùng mà Word có thể dùng để tìm kiếm các câu hoặc đoạn văn cụ thể.
        
        Yêu cầu tìm kiếm: "${naturalQuery}"

        ĐỊNH DẠNG ĐẦU RA PHẢI LÀ JSON NGUYÊN BẢN (KHÔNG CÓ KÝ TỰ MỞ ĐÓNG MÃ CODE):
        {
          "keywords": ["từ khóa 1", "từ khóa phụ 2"],
          "relevant_phrases": ["cụm từ trích dẫn 1", "cụm từ trích dẫn 2"], // Trích xuất các cụm từ liên quan
          "options": {
            "matchWholeWord": false 
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
        console.log("Prompt tóm tắt:");
        res.json(searchPlan);
    } catch (error) {
        console.error("Lỗi khi gọi Gemini:", error);
        res.status(500).json({ error: "Lỗi nội bộ hoặc lỗi phân tích JSON." });
    }
});

app.post('/api/summarize', async (req, res) => {
    const { text } = req.body;

    if (!text) {
        return res.status(400).json({ error: "Thiếu 'text' để tóm tắt." });
    }

    const prompt = `
        Tóm tắt đoạn văn bản sau bằng tiếng Việt. 
        Đoạn văn bản: "${text}"
    `;
    console.log("Prompt tóm tắt:", prompt);

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash', 
            contents: [{ role: "user", parts: [{ text: prompt }] }],
            config: {
                temperature: 0.2
            }
        });

        res.json({ summary: response.text.trim() });
    } catch (error) {
        console.error("Lỗi khi tóm tắt bằng Gemini:", error);
        res.status(500).json({ error: "Lỗi nội bộ khi gọi API tóm tắt." });
    }
});

app.listen(port, () => {
    console.log(`Backend server đang chạy tại http://localhost:${port}`);
});