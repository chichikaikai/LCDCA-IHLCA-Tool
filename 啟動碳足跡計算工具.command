#!/bin/bash

# 切換到這個檔案所在的資料夾
cd "$(dirname "$0")"

echo "========================================="
echo " IHLCA 計算工具（Integrated Hybrid LCA）"
echo "========================================="
echo ""

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null; then
    osascript -e 'display alert "找不到 Python 3" message "請先安裝 Python 3，前往 https://www.python.org 下載。" as critical'
    exit 1
fi

# 第一次執行時安裝套件
if [ ! -f ".installed" ]; then
    echo "首次執行，安裝必要套件（約需 1-2 分鐘）..."
    pip3 install -r requirements.txt --quiet
    if [ $? -ne 0 ]; then
        osascript -e 'display alert "套件安裝失敗" message "請確認網路連線後再試一次。" as critical'
        exit 1
    fi
    touch .installed
    echo "套件安裝完成！"
fi

# 找一個沒被佔用的 port（從 8501 開始往上試）
PORT=""
for p in 8501 8502 8503 8504 8505; do
    if ! lsof -Pi :$p -sTCP:LISTEN -t >/dev/null 2>&1; then
        PORT=$p
        break
    fi
done

if [ -z "$PORT" ]; then
    echo "8501~8505 都被佔用，嘗試關閉舊 streamlit 程序..."
    pkill -f "streamlit run" 2>/dev/null
    sleep 2
    PORT=8501
fi

echo "啟動中，使用 port $PORT，請稍候..."
echo "（關閉此視窗即可停止程式）"
echo ""

# 5 秒後自動開瀏覽器
(sleep 5 && open "http://localhost:$PORT") &

# 啟動 Streamlit（不加 headless，讓 streamlit 內建瀏覽器邏輯正常）
python3 -m streamlit run app.py \
    --browser.gatherUsageStats false \
    --server.port $PORT
