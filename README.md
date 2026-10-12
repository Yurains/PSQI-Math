# PSQI 睡眠品質計算機

依匹茲堡睡眠品質量表（PSQI）填答，計算七個構面及 0–21 分的總分，並匯出 Excel。

## 網頁版

直接開啟：<https://psqi-calculator-tw.peppyluck6.chatgpt.site/>

網頁版可供取得連結的使用者填答、計分與下載 `.xlsx`。每次下載包含目前這一筆填答；計算及 Excel 產生都在瀏覽器中進行。

## 桌面版

需要 Python 3、PyQt5 與 openpyxl。在本資料夾執行：

```powershell
py -m pip install -r requirements.txt
py psqi_calculator.py
```

填答後按「計算分數」檢視結果，或按「匯出 Excel」儲存。桌面版可將新資料列加入欄位格式相同的既有 `.xlsx` 檔。

## Excel 欄位

匯出格式對應提供的 30 欄範例。`Date` 在填答時選擇；`ID` 和 `Time` 會留白，供下載後在 Excel 自行填寫。其餘欄位包含題目答案、七個構面分數與 PSQI 總分。第 5j 題可不作答，計分時以 0 分處理。

總分高於 5 分時，介面會顯示「睡眠品質不佳」。
