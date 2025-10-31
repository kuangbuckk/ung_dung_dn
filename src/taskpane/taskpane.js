// Thay đổi nếu backend của bạn chạy trên cổng hoặc host khác
const BACKEND_URL = "http://localhost:3001/api/process-search"; 

// Hàm quản lý trạng thái tải (sử dụng các tham số DOM)
function setSearching(isSearching, button, spinner, statusDiv) {
    if (isSearching) {
        button.disabled = true;
        statusDiv.textContent = 'Đang phân tích yêu cầu bằng Gemini...';
        button.textContent = 'Đang phân tích...';
        // Hiển thị spinner nếu tồn tại
        if (spinner) spinner.classList.remove('d-none'); 
    } else {
        button.disabled = false;
        button.textContent = 'Tìm kiếm bằng Gemini';
        // Ẩn spinner nếu tồn tại
        if (spinner) spinner.classList.add('d-none');
    }
}

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

// Hàm chính xử lý logic tìm kiếm (nhận các tham số DOM)
async function runSearch(naturalQuery, searchButton, statusDiv, resultsList, loadingSpinner) {
    setSearching(true, searchButton, loadingSpinner, statusDiv);
    resultsList.innerHTML = "";
    const queryValue = naturalQuery.value;

    await Word.run(async (context) => {

        let searchPlan;
        try {
            // Gọi Backend API (Gemini)
            const response = await fetch(BACKEND_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ naturalQuery: queryValue })
            });
            searchPlan = await response.json();

            if (searchPlan.error) throw new Error(searchPlan.error);
            
            statusDiv.textContent = `Gemini đã trích xuất: ${searchPlan.keywords.join(", ")}. Đang tìm kiếm...`;

            // Thực hiện tìm kiếm trong Word
            const body = context.document.body;
            const searchOptions = {
                ignorePunct: true,
                matchCase: false,
                matchWholeWord: searchPlan.options.matchWholeWord || false
            };

            let totalResults = 0;
            
            for (const keyword of searchPlan.keywords) {
                const searchResults = body.search(keyword, searchOptions);
                context.load(searchResults, 'items');
                await context.sync();

                if (searchResults.items.length > 0) {
                    totalResults += searchResults.items.length;
                    
                    // 1. Tải thuộc tính 'text' cho TẤT CẢ các đối tượng Range
                    searchResults.items.forEach((range) => {
                        range.font.highlightColor = '#FFFF00'; // Đánh dấu màu vàng
                        context.load(range, 'text'); // Tải thuộc tính text
                    });

                    await context.sync(); // ĐỒNG BỘ để lấy giá trị text đã tải
                    
                    // 2. Bây giờ, xử lý từng kết quả (thuộc tính text đã sẵn sàng)
                    searchResults.items.forEach((range) => {
                        // Lấy giá trị văn bản TRỰC TIẾP từ thuộc tính đã được tải
                        const textContent = range.text; 
                        
                        // Tạo phần tử list item
                        const li = document.createElement('li');
                        li.className = 'search-result-item'; 
                        
                        // Sử dụng textContent
                        li.textContent = `[${keyword}] - "${textContent.substring(0, 50).trim()}..."`;
                        
                        // Sự kiện click vẫn cần tham chiếu đến đối tượng range API
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
        // Luôn tắt trạng thái tìm kiếm
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