紐約花蓮慈惠堂　線上點燈和法會報名系統

檔案
  index.html   報名頁（網址首頁）：點燈報名、法會報名、繳費單
  admin.html   後台（網址 /admin）：搜尋、收款、密碼設定
  vercel.json  讓網址不用加 .html

目前狀態：示範版
  - 報名資料只存在填寫者自己的瀏覽器，堂裡收不到。
  - 後台示範密碼：義工 1234、管理 8888。
  - 接上堂裡的 Google 試算表後，才會正式收資料。

常改的設定（都在檔案最上面的 CONFIG）
  index.html：法會場次（events，open:true 才會顯示）、價格、燈種、電話
  admin.html：法會場次名稱（和 index.html 一致）
