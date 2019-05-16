# Simulation-Curriculum （SE Project）
# （以前軟工的作業）

1. 導入jar包的方法:
http://blog.csdn.net/mazhaojuan/article/details/21403717 其中第三個

2. 要注意改地址的以 // @@@ 標記了

3. 已Apache Poi讀excel參考
http://blog.csdn.net/wangjianyu0115/article/details/51344853
http://blog.csdn.net/joyous/article/details/51115899
http://www.voidcn.com/blog/u011159417/article/p-6332101.html

4. main 要加throws FileNotFoundException, FileFormatException, IOException 

5. 模擬選課原設定是獨立頁面，所以除非有用多執行，否則進行模擬選課時，麻煩不要顯示其他東西

6. 呼叫需傳入("科系縮寫名", "學號", "目前哪個學期(1 2 3 4 5 6 7 8)")，但請注意第7點目前可用的案例

7. 目前已建立檔案
學號 : 都可
科系 : IECS
學期 : 下學期(2 4 6 8)

8. 模擬選課完以後，當學期所選的名稱和代碼，傳回Course的Name和Code

9. 目前因為搞不懂定位而相當於沒有繼承Course...完全沒用到Course裡的東西

10. 以.Switch進入模擬課表

11. 結束模擬課表後，會回傳字串陣列，目前想法為存到所建的Course，以couse.setCourseName進入，但其實courseName和courseCode都設定了

-----------------------------------------------
尚未處理 : 
必修的處理 一開始的學分載入 // 未有相應檔案可測試 建好後補上


差太多:
查詢歷史
