/* global Office, document */

function log(msg) {
    const consoleDiv = document.getElementById("debug-console");
    if (consoleDiv) {
        const time = new Date().toLocaleTimeString();
        consoleDiv.innerHTML += `[${time}] ${msg}<br>`;
        consoleDiv.scrollTop = consoleDiv.scrollHeight;
    }
}

let pullTimer;

Office.onReady(() => {
    log("UI Ready. Starting PULL request...");

    // 1. 註冊接收器
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onParentMessageReceived
    );

    // 2. 【關鍵】主動向 Parent 要資料 (每秒一次)
    pullTimer = setInterval(() => {
        try {
            Office.context.ui.messageParent("PULL_DATA");
            // log("Sent: PULL_DATA"); // 怕洗版可以註解掉
        } catch (e) {
            log("Wait...");
        }
    }, 1000);

    // 立即先要一次
    Office.context.ui.messageParent("PULL_DATA");

    // 按鈕綁定
    document.getElementById("btnSend").onclick = () => {
        log("Sending VERIFIED_PASS...");
        
        // 傳送訊號給 Parent (commands.js)
        Office.context.ui.messageParent("VERIFIED_PASS");
        
        // 【新增】視覺回饋，因為視窗關閉需要一點時間
        const btn = document.getElementById("btnSend");
        btn.innerText = "驗證完成，視窗關閉中...";
        btn.disabled = true;
    };
    
    document.getElementById("btnCancel").onclick = () => {
        Office.context.ui.messageParent("CANCEL");
    };
});

// 當收到 Parent 的回覆
function onParentMessageReceived(arg) {
    try {
        const message = arg.message;
        const data = JSON.parse(message); 
        
        // 情況 A: Parent 還在忙
        if (data.status === "LOADING") {
            log("⏳ Parent is fetching data...");
            return;
        }
        
        // 情況 B: 收到錯誤
        if (data.error) {
            log("❌ Parent Error: " + data.error);
            if(pullTimer) clearInterval(pullTimer);
            return;
        }

        // 情況 C: 收到真正的資料
        if (data.recipients) {
             log("✅ Data Received! Stopping PULL.");
             
             // 停止請求
             if(pullTimer) clearInterval(pullTimer);
             
             // 渲染畫面
             renderData(data);
        }
    } catch (e) {
        log("Error: " + e.message);
    }
}

function renderData(data) {
    const container = document.getElementById("recipients-list");
    container.innerHTML = "";
    
    if (data.recipients && data.recipients.length > 0) {
        data.recipients.forEach((p, i) => {
            const d = document.createElement("div");
            d.className = "item-row";
            d.innerHTML = `
                <input type='checkbox' checked class='verify-check' id='r_${i}' onchange='checkAllChecked()'>
                <label for='r_${i}'>${p.displayName || p.emailAddress}</label>
            `;
            container.appendChild(d);
        });
    } else {
        container.innerHTML = "無收件人";
    }
    
    // 附件
    const attContainer = document.getElementById("attachments-list");
    attContainer.innerHTML = "";
    if (data.attachments && data.attachments.length > 0) {
        data.attachments.forEach((a, i) => {
            const d = document.createElement("div");
            d.className = "item-row";
            d.innerHTML = `
                <input type='checkbox' checked class='verify-check' id='a_${i}' onchange='checkAllChecked()'>
                <label for='a_${i}'>📎 ${a.name}</label>
            `;
            attContainer.appendChild(d);
        });
    } else {
        attContainer.innerText = "無附件";
    }

    checkAllChecked();
}

window.checkAllChecked = function() {
    const all = document.querySelectorAll(".verify-check");
    let pass = true;
    all.forEach(c => { if(!c.checked) pass = false; });
    
    const btn = document.getElementById("btnSend");
    if (all.length === 0) pass = true;
    
    btn.disabled = !pass;
    if (pass) {
        btn.style.opacity = "1";
        btn.style.cursor = "pointer";
    } else {
        btn.style.opacity = "0.5";
        btn.style.cursor = "not-allowed";
    }
};