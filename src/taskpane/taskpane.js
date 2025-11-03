// taskpane.js
const BACKEND_URL = "http://localhost:3001/api"; 

// =======================================================
// A. HÀM CHUNG VÀ CƠ SỞ
// =======================================================

// Hàm quản lý trạng thái tải (tổng quát)
function setProcessing(isProcessing, button, statusDiv, processName) {
    const spinner = document.getElementById("loading-spinner");
    if (isProcessing) {
        button.disabled = true;
        statusDiv.textContent = `Đang thực hiện ${processName} bằng Gemini...`;
        button.textContent = `Đang xử lý...`;
        if (spinner) spinner.classList.remove('d-none'); 
    } else {
        button.disabled = false;
        if (spinner) spinner.classList.add('d-none');
        // Đặt lại tên nút (cần cập nhật cho phù hợp với 2 tính năng mới)
        if (processName === 'Q&A') button.textContent = 'Trả lời Câu hỏi';
        else if (processName === 'Giải thích') button.textContent = 'Giải thích Thuật ngữ (Kèm Nghiên Cứu)';
    }
}

// Hàm lấy văn bản được chọn
async function getSelectedText(context) {
    const range = context.document.getSelection();
    context.load(range, 'text');
    await context.sync();
    return range.text.trim();
}

// Hàm hiển thị kết quả AI
function displayResult(resultsDiv, title, content) {
    resultsDiv.innerHTML = ''; 
    
    const h4 = document.createElement('h4');
    h4.textContent = title;
    h4.style.borderBottom = '1px solid #ccc';
    h4.style.paddingBottom = '5px';
    resultsDiv.appendChild(h4);

    const p = document.createElement('p');
    p.textContent = content; 
    p.style.whiteSpace = 'pre-wrap'; 
    resultsDiv.appendChild(p);
}


// =======================================================
// B. LOGIC GỌI API (KHÔNG CÓ WORD.RUN)
// =======================================================

async function runQNALogic(documentText, userQuestion, resultsDiv) {
    const response = await fetch(`${BACKEND_URL}/qna`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ documentText, userQuestion })
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: `Lỗi HTTP: ${response.status}` }));
        throw new Error(`Lỗi Server (${response.status}): ${errorData.error || 'Phản hồi không thành công'}`);
    }
    
    const data = await response.json();
    if (data.error) throw new Error(data.error);

    displayResult(resultsDiv, `Trả lời cho: "${userQuestion}"`, data.result);
}

async function runExplainLogic(term, resultsDiv) {
    const response = await fetch(`${BACKEND_URL}/explain`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ term })
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: `Lỗi HTTP: ${response.status}` }));
        throw new Error(`Lỗi Server (${response.status}): ${errorData.error || 'Phản hồi không thành công'}`);
    }
    
    const data = await response.json();
    if (data.error) throw new Error(data.error);

    displayResult(resultsDiv, `Giải thích Thuật ngữ: "${term}"`, data.result);
}


// =======================================================
// C. HÀM WRAPPER (CHỨA WORD.RUN VÀ LOGIC CHUNG)
// =======================================================

// Wrapper cho Hỏi & Đáp Theo Ngữ Cảnh
async function runQNA(statusDiv, resultsDiv, button, questionInput) {
    const processName = 'Q&A';
    console.log(`Bắt đầu quy trình: ${processName}`);
    setProcessing(true, button, statusDiv, processName);

    await Word.run(async (context) => {
        try {
            const documentText = await getSelectedText(context);
            const userQuestion = questionInput.value.trim();

            if (!documentText) throw new Error('Vui lòng chọn văn bản (ngữ cảnh) để hỏi.');
            if (!userQuestion) throw new Error('Vui lòng nhập câu hỏi.');

            await runQNALogic(documentText, userQuestion, resultsDiv);
            statusDiv.textContent = 'Hoàn tất Q&A.';

        } catch (error) {
            statusDiv.textContent = `LỖI ${processName}: ${error.message}`;
            console.error(`Lỗi ${processName}: `, error);
        }
    }).finally(() => {
        setProcessing(false, button, statusDiv, processName);
    });
}

// Wrapper cho Giải thích Thuật ngữ
async function runExplain(statusDiv, resultsDiv, button) {
    const processName = 'Giải thích';
    console.log(`Bắt đầu quy trình: ${processName}`);
    setProcessing(true, button, statusDiv, processName);

    await Word.run(async (context) => {
        try {
            const term = await getSelectedText(context);
            if (!term) throw new Error('Vui lòng chọn thuật ngữ cần giải thích.');
            
            statusDiv.textContent = `Đang tìm kiếm và giải thích thuật ngữ "${term}"...`;

            await runExplainLogic(term, resultsDiv);
            statusDiv.textContent = 'Hoàn tất Giải thích.';

        } catch (error) {
            statusDiv.textContent = `LỖI ${processName}: ${error.message}`;
            console.error(`Lỗi ${processName}: `, error);
        }
    }).finally(() => {
        setProcessing(false, button, statusDiv, processName);
    });
}


// =======================================================
// D. KHỞI TẠO VÀ GÁN SỰ KIỆN
// =======================================================
Office.onReady(info => {
    console.log("Office Ready! Host Type:", info.host);
    if (info.host === Office.HostType.Word) {
        
        // Lấy các phần tử DOM chung
        const statusDiv = document.getElementById("status");
        const resultsDiv = document.getElementById("results-content"); 

        // TÍNH NĂNG ĐƯỢC GIỮ LẠI:
        const qnaInput = document.getElementById("qna-input");
        const qnaButton = document.getElementById("qna-button");
        const explainButton = document.getElementById("explain-button");
        
        // Gán sự kiện click
        if (qnaButton && qnaInput) {
            qnaButton.onclick = () => runQNA(statusDiv, resultsDiv, qnaButton, qnaInput);
        }
        if (explainButton) {
            explainButton.onclick = () => runExplain(statusDiv, resultsDiv, explainButton);
        }

        statusDiv.textContent = 'Sẵn sàng. Vui lòng chọn văn bản và sử dụng các tính năng.';
    }
});