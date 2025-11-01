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
        // Đặt lại tên nút (cần cải tiến nếu có nhiều nút hơn)
        if (processName === 'Tóm tắt') button.textContent = 'Tóm tắt Văn bản Đã Chọn';
        else if (processName === 'Q&A') button.textContent = 'Trả lời Câu hỏi';
        else if (processName === 'Giải thích') button.textContent = 'Giải thích Thuật ngữ (Kèm Nghiên Cứu)';
        else if (processName === 'Tìm kiếm NLP') button.textContent = 'Tìm kiếm bằng Gemini';
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

// Hàm cuộn đến vị trí Range được tìm thấy
function scrollToRange(range) {
    Word.run(async (context) => {
        range.select("Start"); 
        await context.sync();
    });
}


// =======================================================
// B. LOGIC GỌI API (KHÔNG CÓ WORD.RUN)
// =======================================================

async function runSummarizeLogic(documentText, resultsDiv) {
    const summaryRequest = 'Tóm tắt các ý chính trong 3-5 gạch đầu dòng.'; 
    const response = await fetch(`${BACKEND_URL}/summarize`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ documentText, summaryRequest })
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: `Lỗi HTTP: ${response.status}` }));
        throw new Error(`Lỗi Server (${response.status}): ${errorData.error || 'Phản hồi không thành công'}`);
    }
    
    const data = await response.json();
    if (data.error) throw new Error(data.error);
    
    displayResult(resultsDiv, 'Tóm tắt Văn bản:', data.result);
}

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

// Wrapper cho Tóm tắt Văn bản
async function runSummarize(statusDiv, resultsDiv, button) {
    const processName = 'Tóm tắt';
    console.log(`Bắt đầu quy trình: ${processName}`);
    setProcessing(true, button, statusDiv, processName);
    
    await Word.run(async (context) => {
        try {
            const documentText = await getSelectedText(context);
            if (!documentText) throw new Error('Vui lòng chọn văn bản cần tóm tắt.');
            
            await runSummarizeLogic(documentText, resultsDiv);
            statusDiv.textContent = 'Hoàn tất Tóm tắt.';

        } catch (error) {
            statusDiv.textContent = `LỖI ${processName}: ${error.message}`;
            console.error(`Lỗi ${processName}: `, error);
        }
    }).finally(() => {
        setProcessing(false, button, statusDiv, processName);
    });
}

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

// Wrapper cho Tìm kiếm NLP (TÍNH NĂNG CŨ)
async function runSearch(naturalQuery, searchButton, statusDiv, resultsList) {
    const processName = 'Tìm kiếm NLP';
    console.log(`Bắt đầu quy trình: ${processName}`);
    setProcessing(true, searchButton, statusDiv, processName);
    resultsList.innerHTML = "";
    const queryValue = naturalQuery.value;

    await Word.run(async (context) => {
        let searchPlan;
        try {
            // Gọi Backend API (Gemini)
            const response = await fetch(`${BACKEND_URL}/process-search`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ naturalQuery: queryValue })
            });
            
            if (!response.ok) {
                 throw new Error(`Lỗi HTTP khi gọi NLP Search: ${response.status}`);
            }
            searchPlan = await response.json();

            if (searchPlan.error) throw new Error(searchPlan.error);
            
            statusDiv.textContent = `Gemini đã trích xuất: ${searchPlan.keywords.join(", ")}. Đang tìm kiếm...`;

            // Thực hiện tìm kiếm trong Word (Logic cũ của bạn)
            const body = context.document.body;
            const searchOptions = {
                ignorePunct: true, matchCase: false,
                matchWholeWord: searchPlan.options.matchWholeWord || false
            };

            let totalResults = 0;
            
            for (const keyword of searchPlan.keywords) {
                const searchResults = body.search(keyword, searchOptions);
                context.load(searchResults, 'items');
                await context.sync();

                if (searchResults.items.length > 0) {
                    totalResults += searchResults.items.length;
                    
                    searchResults.items.forEach((range) => {
                        range.font.highlightColor = '#FFFF00'; 
                        context.load(range, 'text'); 
                    });

                    await context.sync(); 
                    
                    searchResults.items.forEach((range) => {
                        const textContent = range.text; 
                        const li = document.createElement('li');
                        li.className = 'search-result-item'; 
                        li.textContent = `[${keyword}] - "${textContent.substring(0, 50).trim()}..."`;
                        li.onclick = () => scrollToRange(range);
                        resultsList.appendChild(li);
                    });
                }
            }
            
            statusDiv.textContent = `Hoàn tất tìm kiếm. Tìm thấy ${totalResults} kết quả.`;

        } catch (error) {
            statusDiv.textContent = `LỖI TÌM KIẾM: ${error.message}.`;
            console.error("Lỗi: ", error);
        }
        
    }).catch(error => {
        statusDiv.textContent = `LỖI WORD API: ${error.message}`;
        console.error("Lỗi Word API: " + error.message);
    }).finally(() => {
        setProcessing(false, searchButton, statusDiv, processName);
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

        // TÍNH NĂNG MỚI:
        const summarizeButton = document.getElementById("summarize-button");
        const qnaInput = document.getElementById("qna-input");
        const qnaButton = document.getElementById("qna-button");
        const explainButton = document.getElementById("explain-button");
        
        // TÍNH NĂNG CŨ:
        const naturalQuery = document.getElementById("natural-query");
        const searchButton = document.getElementById("search-button");
        const resultsList = document.getElementById("results-list");

        // Gán sự kiện click cho tính năng mới
        if (summarizeButton) {
            summarizeButton.onclick = () => runSummarize(statusDiv, resultsDiv, summarizeButton);
        }
        if (qnaButton && qnaInput) {
            qnaButton.onclick = () => runQNA(statusDiv, resultsDiv, qnaButton, qnaInput);
        }
        if (explainButton) {
            explainButton.onclick = () => runExplain(statusDiv, resultsDiv, explainButton);
        }
        
        // Gán sự kiện click cho tính năng cũ
        if (searchButton && naturalQuery) {
            searchButton.onclick = () => runSearch(naturalQuery, searchButton, statusDiv, resultsList);
        }

        statusDiv.textContent = 'Sẵn sàng. Vui lòng chọn văn bản và sử dụng các tính năng.';
    }
});