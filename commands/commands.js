/* global Office */

Office.onReady();

// 1. 發送攔截 (OnMessageSend)
function validateSend(event) {
    // 讀取這封信的自訂屬性 'isVerified'
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");

        if (isVerified === true) {
            // A. 已經驗證過了 -> 放行
            // 這裡建議不要移除屬性，避免網路不穩時 Outlook 重試發送會失敗
            // 如果你希望每次修改內容後都要重測，那是更進階的邏輯 (偵測 ItemChanged)，目前先這樣即可
            event.completed({ allowEvent: true });
        } else {
            // B. 還沒驗證 -> 阻擋
            // 系統會自動跳出提示框，顯示下方的 errorMessage
            event.completed({ 
                allowEvent: false, 
                errorMessage: "請點擊上方「開啟檢查清單」按鈕，確認收件人後再發送。" 
            });
        }
    });
}

// 註冊全域函式
if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;