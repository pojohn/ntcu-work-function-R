library(tidyverse)
library(openxlsx)
library(readxl)
library(data.table)
old.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200930-匯入因材網資料轉換/109年計算/正式區"
new.path<-list.files(path = old.path,full.names = T)#指定資料夾內的資料夾清單
test   = "off"
result = "on"

# 10/26 新增資料 因材網 user_id 
# 將 user_id       priori_name  加入 
{
  user.id <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200930-匯入因材網資料轉換/109年計算/學生id/1082學年度學力學生資料.xlsx" ,
                        sheet = 1,
                        col_types = c("text","text","text","text","text",
                                      "text","text","text","text","text"))
  # 去除多餘變數
  user.id <-  user.id[,c(-6)]
  # 轉換編碼
  user.id[["學生名稱"]]   <- enc2utf8(user.id[["學生名稱"]])
  
  # 更新年級格式 +1 與  "0"
  user.id[,2] <- user.id[,2] %>% mutate_if(is.character, as.numeric) 
  user.id[,2] <- user.id[,2]+1
  user.id[,2] <- gsub("^" ,  paste( "0",  sep=""),user.id[["年級"]] )
  
  
  
  
}
# 注意事項: 能力指標的年級分頁有+1，作答反應的檔名與內容沒有+1，但是有在程式內+1
# 1:length(new.path)
system.time(
for (i in 1:length(new.path) ){

  # 設定 單一分頁 筆數 
  person.n <- 15000 
  # 讀入 科目 * 年級 檔案   
  # df.BR.person
  {
    df.original   <- fread(new.path[i],
                           colClasses=list(character=1:6),  # 文字格式 才能保留 學校ID 開頭的 "0"
                           na.strings = "", encoding = "UTF-8") #  備用 encoding = "UTF-8"
    # 刪除 多餘的 "年級" 欄位
    df.original   <- df.original[,-2]



    # 整理檔案格式 與 所需欄位
    {
      # 新增內文標題 (左上) 
      ifelse(df.original[["年級"]][1] <6 , 
             Body.title  <-  paste( "補救教學評量系統 - 202006 ",df.original[["科目名稱"]][1]," - 109年國小基本學力檢測報告統計表",  sep="") ,
             Body.title  <-  paste( "補救教學評量系統 - 202006 ",df.original[["科目名稱"]][1]," - 109年國中基本學力檢測報告統計表",  sep="") )
      

      # 更新年級格式 +1 與  "0"
      df.year.old <- df.original[["年級"]][1]
      df.year.new <- gsub(df.year.old ,  paste( "0",(df.year.old+1),  sep="") ,df.year.old)
      df.original[["年級"]] <- gsub(df.year.old,  df.year.new ,df.original[["年級"]])
      
      
      # 更新 班級 格式 
      # df.original[["班級"]]<-substr(df.original[["班級"]], 2, 3)
      df.original[["班級"]] <- ifelse(df.original[["班級"]] %>% nchar() == 3 ,
                               substr(df.original[["班級"]], 2, 3) , 
                               gsub("^" ,  paste( "0",  sep=""),df.original[["班級"]] )  ) 
      df.original[["班級"]] <- ifelse(df.original[["班級"]] %>% nchar() == 3 ,
                               substr(df.original[["班級"]], 2, 3) , 
                               substr(df.original[["班級"]], 1, 2)) 
      
      # 更新 座號 格式
      df.original[["座號"]] <- ifelse(df.original[["座號"]] %>% nchar() == 2 ,
                               substr(df.original[["座號"]], 1, 2) , 
                               gsub("^" ,  paste( "0",  sep=""),df.original[["座號"]] )  )
      

      

      # ,"學生名稱"
      df.original<-left_join(df.original,user.id, by=c("學校ID","年級","班級","座號","學生名稱"))
      # 合併 學生 個人資訊
      df.original.1 <- df.original[,1:10] %>% unite(學生資訊, 學校ID,年級, 班級,座號,學生名稱, sep = "")
    }


    
    
    
    
    # 各二元作答反應
    BR.name.length   <- strsplit(c(df.original[["二元作答反應"]]) ,  "@XX@")[[1]]  %>% length()
    BR.name.original <- seq(from = 1, to = BR.name.length, by=1)
    BR.name.original <- gsub("^{1}", paste( "第",  sep=""),BR.name.original)
    
    df.BR   <- df.original[,14] %>% separate(`二元作答反應`
                                             , c(BR.name.original), sep = "@XX@") %>% data.frame()
    # 將 空值 部分處理為 0
    df.BR[df.BR  == ""]  <- 0
    # 將 各題 轉換為 數值後 矩陣化
    df.BR.person <- df.BR %>% mutate_if(is.character, as.numeric) %>% as.matrix()
    
    # 將 0 的部分 處理為 1 
    df.BR[df.BR == 0] <- 1
    df.BR.allone <- df.BR %>% mutate_if(is.character, as.numeric) %>% as.matrix()

    
    # 替換 年級
    df.year.out <- (df.year.old+1)
    df.year.out <- gsub("^1$",  paste( "一",  sep=""),df.year.out )
    df.year.out <- gsub("^2$",  paste( "二",  sep=""),df.year.out )
    df.year.out <- gsub("^3$",  paste( "三",  sep=""),df.year.out )
    df.year.out <- gsub("^4$",  paste( "四",  sep=""),df.year.out )
    df.year.out <- gsub("^5$",  paste( "五",  sep=""),df.year.out )
    df.year.out <- gsub("^6$",  paste( "六",  sep=""),df.year.out )
    df.year.out <- gsub("^7$",  paste( "七",  sep=""),df.year.out )
    df.year.out <- gsub("^8$",  paste( "八",  sep=""),df.year.out )
    df.year.out <- gsub("^9$",  paste( "九",  sep=""),df.year.out )
    
    
    匯出年級 <- paste( df.original[["科目名稱"]][1],df.year.out ,  sep="")
    # 轉化 科目 代號
    df.original[["科目名稱"]] <- gsub("自然$"  ,  paste( "S",  sep=""),df.original[["科目名稱"]])
    df.original[["科目名稱"]] <- gsub("數學$"  ,  paste( "M",  sep=""),df.original[["科目名稱"]])
    df.original[["科目名稱"]] <- gsub("國語文$",  paste( "C",  sep=""),df.original[["科目名稱"]])
    df.original[["科目名稱"]] <- gsub("英語文$",  paste( "E",  sep=""),df.original[["科目名稱"]])
    指定年級 <- paste( df.original[["科目名稱"]][1],(df.year.old+1) ,  sep="")
  }
  
  # 指定 答案
  # ans.use.W
  {
    ans.original <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200930-匯入因材網資料轉換/109答對率分布-1027.xlsx" ,
                               sheet = 指定年級,
                               col_types = c("numeric","text","text","text","text"))
    ans.original[["題號"]] <- gsub("^{1}", paste( "第",  sep=""),ans.original[["題號"]])
    ans.use      <- ans.original[,c("代號","題號","能力指標")]  
    
    # 將 各向度 長轉寬 表格 
    ans.use.W <- pivot_wider(ans.use,names_from = 能力指標 , values_from = 代號)
    # 將 NA部分處理為 0
    ans.use.W[is.na(ans.use.W)  == TRUE]  <- 0
    ans.use.W.out <- ans.use.W
    ans.use.W <- as.matrix(ans.use.W[,-1])
  }  
  
  
  # 計算 與 整理 資料
  {
    # 計算 各 能力指標 的 個人 答對數
    df.BR.person.ans   <- df.BR.person %*% ans.use.W 
    # 計算 各 能力指標 的 all one  答對數
    df.BR.allone.ans   <- df.BR.allone %*% ans.use.W 
    # 核對 答對數 是否相同 ，相同為 1，不相同為 0
    df.BR.end.ans      <- ifelse( df.BR.person.ans == df.BR.allone.ans , 1, 0) %>% as.data.table()
    
    # 將 0 1 轉換為 O X
    df.new.ans <- data.table()
    
    df.new.ans <- ifelse( df.BR.end.ans == 1 , "○", "X")
    
    # 合併 學生 資訊 欄位 
    df.end.ans         <- data.table(
                                     學生資訊    = df.original.1[,學生資訊],
                                     user_id     = df.original[,user_id],
                                     priori_name = df.original[,priori_name],
                                     學校ID = df.original[,學校ID],
                                     年級   = df.original[,年級],
                                     班級   = df.original[,班級],
                                     座號   = df.original[,座號],
                                     df.new.ans)
    # 轉換編碼
    names(df.end.ans)   <- enc2utf8(names(df.end.ans))
    
    
    df.end.ans.c <- df.end.ans %>% mutate_if(is.numeric, as.character) 
    
    # 將 各能力指標 寬轉長 表格
    df.end.ans.L <- pivot_longer(df.end.ans.c, cols = - 學生資訊, names_to = "能力指標", values_to = "代碼"  )
    # 將 各能力指標 長轉寬 表格
    df.end.ans.W <- pivot_wider (df.end.ans.L ,names_from = 學生資訊 , values_from = 代碼)
    # 計算 要分成 幾個分檔
    df.cut.n  <- ceiling(ncol(df.end.ans.W)/person.n)
    # df.end.ccc   <-split(df.end.ans.L , cut(seq(df.end.ans.W), nrow(df.end.ans.W), labels = FALSE))
    
    # 計算二元作答反應 答對答錯總數
    df.BR.end.ans.person   <- data.table(學生資訊 = df.original.1[,學生資訊], df.BR.end.ans)
    df.BR.end.ans.pL  <- pivot_longer(df.BR.end.ans.person, cols = - 學生資訊, names_to = "能力指標", values_to = "代碼"  )
    df.BR.end.ans.pW  <- pivot_wider (df.BR.end.ans.pL ,names_from = 學生資訊 , values_from = 代碼)

    
  }
  
  # 核對用
  if( test == "on")  {
    fwrite(df.BR.person,
           file = paste( "C:/Users/user/Desktop/測試/測試",指定年級,"個人作答",".csv" ,  sep=""),
           sep  = ",",
           bom  = T )   
    
    fwrite(ans.use.W.out,
           file = paste( "C:/Users/user/Desktop/測試/測試",指定年級,"能力指標",".csv" ,  sep=""),
           sep  = ",",
           bom  = T ) 
    
    
    fwrite(df.BR.end.ans,
           file = paste( "C:/Users/user/Desktop/測試/測試",指定年級,"核對結果",".csv" ,  sep=""),
           sep  = ",",
           bom  = T )
  }
  # 結果用
  if( result == "on"){
    # 匯出 資料
    for (j in 1:df.cut.n){
      # O、X 版本 (df.out.right)
      if(j == 1 & ncol(df.end.ans.W) < person.n){ df.out.right <- df.end.ans.W[,c(1:(ncol(df.end.ans.W)))] }  else{  df.out.right <- df.end.ans.W[,c(1:person.n)] }
      if(j>1 & j <  df.cut.n ){ df.out.right <- df.end.ans.W[,c(1,((j-1)*person.n+1):((j)*person.n))]}
      if(j>1 & j == df.cut.n ){ df.out.right <- df.end.ans.W[,c(1,((j-1)*person.n+1):(ncol(df.end.ans.W)))] }
      #  1、0 版本 (df.out.left)
      if(j == 1 & ncol(df.BR.end.ans.pW) < person.n){ df.out.left <- df.BR.end.ans.pW[,c(1:(ncol(df.BR.end.ans.pW)))] }  else{  df.out.left <- df.BR.end.ans.pW[,c(1:person.n)] }
      if(j>1 & j <  df.cut.n ){ df.out.left <- df.BR.end.ans.pW[,c(1,((j-1)*person.n+1):((j)*person.n))]}
      if(j>1 & j == df.cut.n ){ df.out.left <- df.BR.end.ans.pW[,c(1,((j-1)*person.n+1):(ncol(df.BR.end.ans.pW)))] }
      left.1 <- df.out.left[,-1] %>% rowSums()
      left.sum <- df.out.left[,-1] %>% length()
      left.0 <- left.sum -  left.1
      df.out.left.end <- data.table("序號"         = seq(1:length(left.1)),
                                    "基本學習內容" = df.out.left[[1]],
                                    "能力指標"     = df.out.left[[1]],
                                    "○"            = left.1 ,
                                    "△"            = "0" ,
                                    "X"            = left.0,
                                    "合計"         = left.sum)
      

      #開啟 新的 Workbook
      wb <- createWorkbook()
      # 設定 文件 字體、顏色、大小
      modifyBaseFont(wb, fontSize = 12, fontColour = "#000000", fontName = "新細明體")
      # 新增分頁
      addWorksheet(wb, sheetName = paste( 指定年級,"-",j ,  sep=""))   
      
      #part1  : Body.title
      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 1,
                startRow = 1,
                x = Body.title) 
      
      #part2  : df.out.left.end

      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 1,
                startRow = 8,
                x = df.out.left.end) 
      # 補上合併後的欄位名稱
      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 1,
                startRow = 2,
                x = df.out.left.end[1,1:3] ) 
      
      
      
      #part3  : df.out.right

      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 8,
                startRow = 2,
                x = df.out.right[,-1]) # headerStyle = hs1 ,borders = "all",
      # 移動 欄位
      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 8,
                startRow = 1,
                x = df.out.right[1:6,-1])
      #  刪除多餘資料
      deleteData(wb, sheet = paste( 指定年級,"-",j ,  sep=""),
                 cols = 8:16000, rows = 1, gridExpand = TRUE)
      # 補上 缺少的 學生資訊
      writeData(wb, sheet = paste( 指定年級,"-",j ,  sep="") ,
                startCol = 8,
                startRow = 8,
                x = df.out.right[0,-1])
      
      #  左上標題 : Merge cells: Row 1 column A to G (1:7)
      mergeCells(wb, paste( 指定年級,"-",j ,  sep=""), cols = 1:7, rows = 1)
      #  左下欄位名稱: Merge cells: Row 2:6 column A & B & C (1,2,3)
      mergeCells(wb, paste( 指定年級,"-",j ,  sep=""), cols = 1, rows = 2:8)
      mergeCells(wb, paste( 指定年級,"-",j ,  sep=""), cols = 2, rows = 2:8)
      mergeCells(wb, paste( 指定年級,"-",j ,  sep=""), cols = 3, rows = 2:8)
      # 匯出資料
      saveWorkbook(wb, paste( "C:/Users/user/Desktop/測試/109檢測報告_", 匯出年級 ,"-",j,".xlsx" ,  sep=""), overwrite = TRUE)

    }
    
  }

}
)
  


