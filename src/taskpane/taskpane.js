// Thay đổi nếu backend của bạn chạy trên cổng hoặc host khác
const BACKEND_URL = "http://localhost:3001/api/process-search"; 

// Khởi tạo Add-in khi Word sẵn sàng
Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        
        // 1. Gán các phần tử DOM cho biến trong phạm vi này
        const naturalQuery = document.getElementById("natural-query");
        const searchButton = document.getElementById("search-button");
        const statusDiv = document.getElementById("status");
        const resultsList = document.getElementById("results-list");
        const loadingSpinner = document.getElementById("loading-spinner"); 

        if (!searchButton || !statusDiv || !resultsList || !naturalQuery) {
             console.error("Lỗi: Không tìm thấy các phần tử HTML cần thiết. Kiểm tra ID trong taskpane.html.");
             statusDiv.textContent = 'Lỗi khởi tạo: Thiếu phần tử HTML.';
             return;
        }

        // 2. Gán sự kiện click, truyền các biến DOM làm tham số
        searchButton.onclick = () => runSearch(naturalQuery, searchButton, statusDiv, resultsList, loadingSpinner);
        
        statusDiv.textContent = 'Sẵn sàng để bắt đầu tìm kiếm.';
    }
});

// Thêm hàm này vào taskpane.js, sau hàm runSearch và scrollToRange
async function summarizeText(textToSummarize) {
    const SUMMARY_URL = "http://localhost:3001/api/summarize"; // Cần tạo endpoint mới
    
    try {
        const response = await fetch(SUMMARY_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: textToSummarize })
        });
        const result = await response.json();
        
        if (result.error) throw new Error(result.error);
        return result.summary;
    } catch (error) {
        console.error("Lỗi tóm tắt Gemini:", error);
        return "Không thể tóm tắt đoạn này.";
    }
}

// Hàm quản lý trạng thái tải (giữ nguyên)
function setSearching(isSearching, button, spinner, statusDiv) {
    if (isSearching) {
        button.disabled = true;
        statusDiv.textContent = 'Đang phân tích yêu cầu bằng Gemini...';
        button.textContent = 'Đang phân tích...';
        if (spinner) spinner.classList.remove('d-none'); 
    } else {
        button.disabled = false;
        button.textContent = 'Tìm kiếm bằng Gemini';
        if (spinner) spinner.classList.add('d-none');
    }
}

// HÀM CHÍNH ĐÃ ĐƯỢC CẬP NHẬT ĐỂ TÌM KIẾM, ĐÁNH DẤU, VÀ TÓM TẮT
async function runSearch(naturalQuery, searchButton, statusDiv, resultsList, loadingSpinner) {
    // Đảm bảo BACKEND_URL được định nghĩa ở đầu file taskpane.js
    const BACKEND_URL = "http://localhost:3001/api/process-search"; 
    
    setSearching(true, searchButton, loadingSpinner, statusDiv);
    resultsList.innerHTML = "";
    const queryValue = naturalQuery.value;

    await Word.run(async (context) => {
        let searchPlan;
        try {
            // Bước 1: Gọi Backend API để lấy keywords và phrases
            const response = await fetch(BACKEND_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ naturalQuery: queryValue })
            });
            searchPlan = await response.json();

            if (searchPlan.error) throw new Error(searchPlan.error);
            
            const searchTerms = [...searchPlan.keywords, ...(searchPlan.relevant_phrases || [])];
            statusDiv.textContent = `Gemini đã trích xuất ${searchTerms.length} cụm từ. Đang tìm kiếm...`;

            // Bước 2: Thực hiện tìm kiếm trong Word
            const body = context.document.body;
            const searchOptions = {
                ignorePunct: true,
                matchCase: false,
                matchWholeWord: searchPlan.options.matchWholeWord || false
            };

            let totalResults = 0;
            const uniqueResults = new Set(); // Dùng để tránh trùng lặp nội dung

            for (const keyword of searchTerms) {
                if (!keyword || keyword.trim() === '') continue;

                const searchResults = body.search(keyword, searchOptions);
                context.load(searchResults, 'items');
                await context.sync();

                if (searchResults.items.length > 0) {
                    
                    // Lấy kết quả khớp đầu tiên cho từ khóa này
                    const firstMatch = searchResults.items[0]; 
                    
                    // Lấy toàn bộ đoạn văn bản chứa kết quả đầu tiên
                    const parentParagraph = firstMatch.getRange("Whole").parentParagraph;
                    
                    // Tải các thuộc tính: Text và Vị trí (Page)
                    context.load(parentParagraph, 'text'); 
                    context.load(firstMatch, 'page'); 
                    context.load(firstMatch, 'text'); // Tải text của range để đánh dấu
                    
                    await context.sync(); // ĐỒNG BỘ để lấy giá trị đã tải

                    const fullParagraphText = parentParagraph.text.trim();
                    const resultKey = fullParagraphText.substring(0, 100); // Dùng 100 ký tự đầu để so sánh

                    if (!uniqueResults.has(resultKey)) {
                        uniqueResults.add(resultKey);
                        totalResults++;
                        
                        // Đánh dấu màu vàng
                        firstMatch.font.highlightColor = '#FFFF00'; 
                        
                        // Bước 3: Gọi Gemini để TÓM TẮT đoạn văn bản
                        statusDiv.textContent = `Đang tóm tắt đoạn văn chứa "${keyword}"...`;
                        const summary = await summarizeText(fullParagraphText);
                        
                        // Lấy thông tin vị trí (số trang)
                        let locationInfo = '';
                        try {
                            const pageNumber = firstMatch.page; 
                            if (pageNumber && pageNumber !== 0) {
                                locationInfo = ` | Trang ${pageNumber}`;
                            }
                        } catch (e) {
                            // Bỏ qua lỗi Page API
                        }
                        
                        // Tạo phần tử list item
                        const li = document.createElement('li');
                        li.className = 'search-result-item'; 
                        
                        // HIỂN THỊ TÓM TẮT
                        li.innerHTML = `
                            <strong>[${keyword}] - Tóm tắt:</strong> ${summary}
                            <br><small style="color:#777;">(${locationInfo})</small>
                        `;
                        
                        // Sự kiện click cuộn đến vị trí
                        li.onclick = () => scrollToRange(firstMatch);
                        resultsList.appendChild(li);
                    }
                }
            }
            
            statusDiv.textContent = `Hoàn tất tìm kiếm. Tìm thấy ${totalResults} kết quả duy nhất.`;

        } catch (error) {
            statusDiv.textContent = `LỖI TÌM KIẾM: ${error.message}. Vui lòng kiểm tra server Node.js hoặc thử lại.`;
            console.error("Lỗi: ", error);
        }
        
    }).catch(error => {
        statusDiv.textContent = `LỖI WORD API: ${error.message}`;
        console.error("Lỗi Word API: " + error.message);
    }).finally(() => {
        setSearching(false, searchButton, loadingSpinner, statusDiv);
    });
}

// Hàm cuộn đến vị trí Range được tìm thấy
function scrollToRange(range) {
    Word.run(async (context) => {
        range.select("Start"); 
        await context.sync();
    });
}