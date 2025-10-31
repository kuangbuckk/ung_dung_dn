const express = require('express');
const { GoogleGenAI } = require('@google/genai');
const cors = require('cors');
require('dotenv').config();

// Lấy khóa API
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "YOUR_FALLBACK_API_KEY"; 
const ai = new GoogleGenAI(GEMINI_API_KEY);

const app = express();
// Đảm bảo cổng này khớp với URL trong Manifest và Frontend
const port = 3000; 

// Cần thiết để Add-in (Frontend) giao tiếp với Backend
app.use(cors()); 
app.use(express.json());

// Endpoint xử lý yêu cầu tìm kiếm NLP
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

        // Xử lý chuỗi JSON trả về
        const jsonText = response.text.trim().replace(/```json|```/g, '');
        const searchPlan = JSON.parse(jsonText);

        res.json(searchPlan);
    } catch (error) {
        console.error("Lỗi khi gọi Gemini:", error);
        res.status(500).json({ error: "Lỗi nội bộ, lỗi phân tích JSON, hoặc API Key không hợp lệ." });
    }
});

app.listen(port, () => {
    console.log(`Backend server đang chạy tại http://localhost:${port}`);
});