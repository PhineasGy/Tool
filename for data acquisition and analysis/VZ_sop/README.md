# VZ sop
Project for: Lixel.inc

Author: Guan Yu Chen (PhineasGy)

Description: 裕度分析完整流程 (實驗，分析，critical)

Version: V1.0

update:

------

(Requirement: BC_alone (class method))

1. Excel: "模板.xlsx"

2. 分析: 使用 BC_Alone 工具，將結果 table 貼上 "分析裕度 (VDXXX)" 工作表

3. 實驗: 在 Module 穩態後進行實驗 (溫度 >= 32 degree)，根據不同尺寸或 Pattern 進行各自所需的實驗，將結果輸入至 "實驗裕度 (Case)" 工作表

4. 實驗裕度整理: 將實驗結果複製貼至 "實驗裕度整理" 工作表，可以觀看趨勢 (optional)

5. 分析裕度整理: 

   將實驗結果複製貼至 "最上面表格"
   
   執行 "derive_failnumber_by_exp.m"，得到 Fail Number Table 和 Critical Median 和 分析裕度(with critical median)

   自行複製到 "分析裕度整理" 工作表 下方兩表格 (optional)

6. 總整理:

   左側表格貼上"實驗裕度"，右側表格貼上"分析裕度 (with critical median)"，中間圖表將顯示兩者差異與趨勢
 
