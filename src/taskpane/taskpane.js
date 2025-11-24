const BACKEND_URL = "http://localhost:3001/api";
let fileId = null;
let fileDisplayName = "Document_Content.txt";

// ===============================
// BACKEND WRAPPER
// ===============================
async function apiPost(path, data) {
    const res = await fetch(`${BACKEND_URL}/${path}`, {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify(data)
    });
    const json = await res.json().catch(()=>({error:`HTTP ${res.status}`}));
    if(!res.ok || json.error) throw new Error(json.error || `HTTP ${res.status}`);
    return json;
}

// ===============================
// UPLOAD LOGIC
// ===============================
async function getFullDocumentText() {
    let fullText = "";
    await Word.run(async context=>{
        const body = context.document.body;
        body.load("text");
        await context.sync();
        fullText = body.text.trim();
    });
    return fullText;
}

function showProgress(percent){
    const container = document.getElementById("upload-progress-container");
    const bar = document.getElementById("upload-progress-bar");
    const text = document.getElementById("upload-progress-text");
    container.classList.remove("d-none");
    bar.style.width = `${percent}%`;
    text.textContent = `${Math.floor(percent)}%`;
}

function hideProgress(){
    const container = document.getElementById("upload-progress-container");
    container.classList.add("d-none");
}

async function runUpload(statusDiv, button){
    button.disabled = true;
    statusDiv.textContent = "Đang upload tài liệu...";
    showProgress(0);

    try{
        const text = await getFullDocumentText();
        if(!text) throw new Error("Tài liệu trống.");

        // Simulate progress
        let progress=0;
        const interval = setInterval(()=>{
            progress += Math.random()*10;
            if(progress>95) progress=95;
            showProgress(progress);
        }, 300);

        // Call backend
        const res = await apiPost("upload-file",{documentText:text, displayName:fileDisplayName});
        clearInterval(interval);
        showProgress(100);
        fileId = res.fileId;
        statusDiv.textContent = `Upload xong (${res.chunks} chunks, ${res.time}s)`;
        console.log(`Indexed ${res.chunks} chunks in ${res.time}s`);

    }catch(e){
        console.error(e);
        statusDiv.textContent = `Lỗi upload: ${e.message}`;
        hideProgress();
    }finally{
        button.disabled=false;
        setTimeout(hideProgress,1000);
    }
}

// ===============================
// CHAT UI
// ===============================
const chatContainer = document.getElementById("chat-container");
const chatInput = document.getElementById("chat-input");
const chatSendBtn = document.getElementById("chat-send-btn");

function addChatMessage(text,sender="bot"){
    const msg = document.createElement("div");
    msg.textContent = text;
    msg.className = sender;
    chatContainer.appendChild(msg);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

async function sendChatQuestion(){
    const q = chatInput.value.trim();
    if(!q) return;
    addChatMessage(q,"user");
    chatInput.value="";
    if(!fileId){ addChatMessage("Vui lòng upload tài liệu trước.","bot"); return; }

    try{
        const data = await apiPost("global-qna",{userQuestion:q});
        addChatMessage(data.result,"bot");

        // Highlight only top 1 citation
        if (data.citations && data.citations.length > 0) {
            const topCitation = data.citations[0];
            await Word.run(async context => {
                const body = context.document.body;
                const highlightTexts = topCitation.text.match(/.{1,200}/g); // chia nhỏ 200 ký tự

                for (const textPart of highlightTexts) {
                    const searchResults = body.search(textPart, { matchCase: false });
                    searchResults.load("items");
                    await context.sync();
                    searchResults.items.forEach(r => r.font.highlightColor = "#FFFF00");
                }

                // Scroll tới first highlight
                if (highlightTexts.length > 0) {
                    const firstResult = body.search(highlightTexts[0], { matchCase: false });
                    firstResult.load("items");
                    await context.sync();
                    if (firstResult.items.length > 0) firstResult.items[0].select();
                }

                await context.sync();

                // Xóa highlight sau 3s
                setTimeout(async () => {
                    await Word.run(async ctx => {
                        const body2 = ctx.document.body;
                        for (const txt of highlightTexts) {
                            const results = body2.search(txt, { matchCase: false });
                            results.load("items");
                            await ctx.sync();
                            results.items.forEach(r => r.font.highlightColor = null);
                        }
                        await ctx.sync();
                    });
                }, 3000);
            });
        }

    }catch(e){
        console.error(e);
        addChatMessage("Lỗi khi gọi server: "+e.message,"bot");
    }
}

chatSendBtn.onclick = sendChatQuestion;
chatInput.addEventListener("keydown",e=>{
    if(e.key==="Enter" && !e.shiftKey){
        e.preventDefault();
        sendChatQuestion();
    }
});

let autoUploaded=false;
Office.onReady(info=>{
    if(info.host!==Office.HostType.Word) return;
    const statusDiv = document.getElementById("status");
    const btnUpload = document.getElementById("index-button");

    if(!autoUploaded){
        autoUploaded=true;
        statusDiv.textContent="Khởi tạo... Đang upload tài liệu...";
        setTimeout(()=>runUpload(statusDiv,btnUpload),500);
    }

    btnUpload.onclick = ()=>runUpload(statusDiv,btnUpload);
});
