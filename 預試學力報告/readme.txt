程式目的
	1.計算成績
	2.產出圖表
文件準備
	1.各科目答案檔(欄位名: 預試題號、評量向度、正確答案)
	2.各科目全體通過率(使用全體通過率(預試) 計算後 加入 各科目答案檔)
	3.各學校題目檔(檔名範例: 縣市-(學校名稱)(科目)、苗栗-六合國小六年級數學)
	  身份欄位設定: 卷別	縣市	學校	班級	座號	姓名	資源班	體育班


程式設定:
  RtoEXCEL程式(第一版0616):
	 school.score.plot.cc(data.path,data,soc.all.in.one,key,plot.turn = F)
	1.data.path: 原始資料路徑(# choose.files())
	2.data :匯入 資料(# read_excel(data.path))
	3.soc.all.in.one:匯入 預設 全體通過率 (# read_excel())
	4.key :匯入 答案 (# key  <- read_excel())
	5.plot.turn:是否需要繪圖
  RtoEXCEL程式(第二版0616):

	school.score.plot.cc.2(data.path,data,soc.all ,plot.turn = F)
	1.data.path: 原始資料路徑(# choose.files())
	2.data :匯入 資料(# read_excel(data.path))
	3.soc.all:匯入 班級 + 校 + 全體通過率 (# read_excel())
	4.plot.turn:是否需要繪圖
狀況:
	1. 有A、B的情況，有需要分別計算出成績後，合併共同欄位再做繪圖，但這樣 第一版 就不適合
	2. 再修改出 能使用 已算好資料的版本 -> RtoEXCEL程式(第二版0616)




匯出檔名範例: 縣市-(學校名稱)(科目)(年度)學力預試報告 -> 苗栗-六合國小六年級數學109年學力預試報告