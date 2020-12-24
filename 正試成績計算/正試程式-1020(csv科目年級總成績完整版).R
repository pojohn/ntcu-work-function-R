library(data.table)
library(tidyverse)
library(openxlsx) # write.xlsx
library(readxl) # read_excel
library(dummies) # dummy
library(rio) # export
library(Rmpfr) # 浮點數
#計算到小數點十五位
options(digits=15)
# "E:\\台中教育大學\\學力工作\\20200729-學力人數統計\\縣市csv合併\\台中市合併.xlsx"
# "C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/縣市csv合併"
old.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/00計分用-0812/00科目csv"
new.path<-list.files(path = old.path,full.names = T)#指定資料夾內的資料夾清單
counties<-new.path %>% basename() %>% strsplit( "[,. ]") 
# length(new.path)
# 迴圈 個別讀檔 加上範圍 
df<-data.table() #   :length(new.path)
system.time(
  
#第一個 迴圈 按照科目讀取檔案 #測試中 先固定一個科目 #  1:4
for(i in 1:4 ){
  # 新縣市 開始前 重置
  df<-data.table()
  dn<-new.path[i] %>% basename() %>% strsplit( "[,. ]") 
  counties.name<-dn[[1]][1]
  # 建立繪圖存檔用 資料夾
  dir.create(paste("C:/Users/user/Desktop/",counties.name,sep = ""))
  # 讀入 檔案
  df <- fread(new.path[i],na.strings = "")
  # 按照 縣市、年級、班級代號、座號 排序 資料
  df <- setorder(df,縣市,年級,班級代碼,座號)

  #班級 必為類別
  df$班級代碼 <- as.factor(df$班級代碼) 
  # 合併檔輸出後 ，將  NA 的 部分 設為 "0"
  df[is.na(df)  == TRUE]  <- "0"
  # 計算 該縣市 科目種類 與 年級 種類
# 科目種類<- df  %>% group_by(測驗科目代碼,年級) %>% summarise(總人數 = n() , .groups = 'drop_last' ) 
  # data.table
  科目種類<- df[,.(N = .N),by = .(測驗科目代碼,年級)]
  # 取得 指定 科目種類 與 年級 種類 作答反應 與 答案
  # 第二個迴圈 按照 科目 與 年級 讀取 資料 與 答案
  # 1:nrow(科目種類)
  for(j in 1:nrow(科目種類) ){
    
    指定科目 <- 科目種類[j,1]
    
    # 指定 答案
    if (指定科目 == "C" )   { 指定年級 <- paste( "國",科目種類[[j,2]] ,  sep="")
    ans <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/答案/109計分用答案檔-國文.xlsx" 
           , sheet = 指定年級 )   }    else{
             if (指定科目 == "E")    {
            指定年級 <- paste( "英",科目種類[[j,2]] ,  sep="")
            ans <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/答案/109計分用答案檔-英文.xlsx" 
                              , sheet = 指定年級 )   }    else{
            if (指定科目 == "M")    {
             指定年級 <- paste( "數",科目種類[[j,2]] ,  sep="")
             ans <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/答案/109計分用答案檔-數學.xlsx" 
             , sheet = 指定年級 )     }  else{  
              if (指定科目 == "S")    {
               指定年級 <- paste( "自",科目種類[[j,2]] ,  sep="")
                ans <- read_excel("C:/Users/user/Documents/台中教育大學/學力工作/20200729-學力人數統計/答案/109計分用答案檔-自然.xlsx" 
               , sheet = 指定年級 )     }
               }
              }
           }
    
    
    # 重置 data.table
    df.s <- data.table()
    
    # 指定縣市 作答資料 ， 排除 缺考的學生
#    df.s <-  df   %>%  
#      filter(測驗科目代碼 == paste( 科目種類[[j,1]] ,  sep="") & 年級 == paste( 科目種類[[j,2]] ,  sep="") & 缺考 == "0") 
    df.s <-  df[測驗科目代碼 == paste( 科目種類[[j,1]] ,  sep="") & 年級 == paste( 科目種類[[j,2]] ,  sep="") & 缺考 == "0"]
    
    # 如果是 數學 四年級 第18題 移除
    if (指定年級 == "數4" ){
      df.s<-df.s[,-39]
    }
    
    # 如果是 數學 五年級 第19題 答案1 跟 2都行 ，改為 9 
    if (指定年級 == "數5" ){
      df.s$第19題<-gsub("1", "9",df.s$第19題)
      df.s$第19題<-gsub("2", "9",df.s$第19題)
      df.s$第19題<-gsub("99", "9",df.s$第19題) #複選1,2也對
    }
    
    # 如果是 數學 六年級 第25題 移除
    if (指定年級 == "數6" ){
      df.s<-df.s[,-46]
    }
    # 如果是 國語文 八年級  第8題 移除
    if (指定年級 == "國8" ){
      df.s<-df.s[,-29]
    }
    # 根據 答案檔 指定題目的數量
    df.s <-  df.s[,1:(21+nrow(ans))]
    
    
    # 特定科目 需要 個別設定
    
    ########################################
    ########################################
    #  成績計算
    ########################################
    ########################################
    # 注意 題目的順序  必須正確(1.2.3.4....n)
    df.q<- df.s[,-1:-21] %>% data.table() #前 21 欄是學生基本資料
    key   <- ans
    # 注意 答案的順序 必須正確(1.2.3.4....n)
    key.a <- c(key$正確答案)  
    
    # 根據 正確答案 核對出 二元作答反應
    files<- length(c(df.s[[1]]))
    df.binary.response <- data.frame()
    key.a.all<- do.call("rbind", rep(list(key.a), files))
    df.binary.response     <- ifelse( df.q == (key.a.all), 1, 0)
    
    
    
    # 將 NA部分處理為 0
    df.binary.response[is.na(df.binary.response)  == TRUE]  <- 0 
    # 修改 二元作答反應欄位名
    names(df.binary.response)<-gsub("第", "二元第",names(df.binary.response))
    
    # 計算 個人通過率、個人平均通過率、個人 各向度通過率
    總平均 <- rowSums(df.binary.response)/ncol(df.binary.response)
    # 分出向度 # 警告訊息 與 計算過程無關
    key.dummy<- data.table(dummy(key$評量向度))
    names(key.dummy)<-c(levels(factor(key$評量向度)))
    
    # 如果是 英文 五年級 第 34 題跨向度
    if (指定年級 == "英5" ){
      key.dummy$`閱讀能力-文化節慶理解`[34] <- "1"
    }
    
    
    
    # 計算向度總分平均
    sets <- apply(key.dummy, 2, function(XXX) data.table(df.binary.response[, XXX == 1]))
    names(sets) <- c(levels(factor(key$評量向度)))
    # 計算向度平均值
    l1<-lapply(sets,function(x)  rowSums(x)/ncol(x))
    d1<-do.call("cbind",l1)
    ##############################################################
    # 個人身分資訊 與 成績 合併
    soc.all<-data.table(df.s[,c(1:9,11,13:21)],總平均,d1) 
    
    # 班 與 校 必須排除 資源班+補考+缺考
    #  全體 的 成績
    soc.total<-soc.all   %>%
      filter(缺考   == "0" & 補考   == "0" & 資源班   == "0")  %>% 
      group_by(測驗科目代碼,年級)     %>%
      summarise_if(is.numeric,~round(mean(.),15), na.rm = FALSE, .groups = 'drop_last')
    
#    soc.total<-soc.all[缺考   == "0" & 補考   == "0" & 資源班   == "0",lapply(.SD, mean),by = .(測驗科目代碼,年級)]
    
    # 計算 學生 PR值(全體含各縣市)
    # dp <- dplyr::percent_rank(soc.all$總平均)
    # soc.rank<- rank(-soc.all$總平均, ties.method ="min") 
    dp.all <- (rank(soc.all$總平均, ties.method = "max", na.last = "keep") -1)/ (sum(!is.na(soc.all$總平均)) )
    soc.rank.all<- rank(-soc.all$總平均, ties.method ="min") 
    
    
    # 無條件捨去用
#   floor_dec  <- function(x, level=1) round(x - 4.999999999999*10^(-level-1), level)

    dp.all <-floor(dp.all*100+0.00000000000001)
    # 將 RP值 = 100 的部分處理為 99
    dp.all[dp.all  == 100]  <- 99
    
    # 根據 縣市 計算 學生 PR值(單獨 各縣市)
    soc.Counties<-soc.all   %>%
      filter(缺考   == "0" )  %>% 
      group_by(縣市,測驗科目代碼,年級)     %>%
      summarise(總人數 = n() , .groups = 'drop_last' ) 
    # 縣市數量
    Counties.files <- nrow(soc.Counties)
    # 縣市名稱
    # Counties.names <- c(levels(factor(soc.Counties$縣市))) 
     Counties.names <- c(unique(soc.all[["縣市"]]))
    
    df.3 <- data.table()
    df.4 <- data.table()
    df.5 <- data.frame()
    df.6 <- data.frame()
    for(k in 1:Counties.files) {
      df.Counties <- soc.all %>% filter(縣市   == Counties.names[k] )
      df.3 <- (rank(df.Counties$總平均, ties.method = "max", na.last = "keep") -1)/ (sum(!is.na(df.Counties$總平均)) )
      df.4 <- rbind(df.4, df.3)  
      df.5 <- rank(-df.Counties$總平均, ties.method ="min") %>% data.frame()
      df.6 <- rbind(df.6, df.5)
    }
    # 將 RP值 = 100 的部分處理為 99
    df.4[df.4  == 100]  <- 99
    names(df.4) = "百分等級(縣市PR值)"
    names(df.6) = "排名(縣市)"
    df.4 <-floor(df.4*100+0.00000000000001)
    # 將 RP值 = 100 的部分處理為 99
    df.4[df.4  == 100]  <- 99
    # 篩選 出 缺考名單
    # df.miss <-  df[,c(-10,-12)] %>%  
    #   filter(測驗科目代碼 == paste( 科目種類[[j,1]] ,  sep="") & 年級 == paste( 科目種類[[j,2]] ,  sep="") & 缺考 == "3") 
    
    # 整合 個人成績
    soc.person <- data.table(df.s[,c(1:9,11,13:21)],df.6,df.4,`排名(全體)` = soc.rank.all,`百分等級(全體PR值)` = dp.all,總平均,d1,df.q,df.binary.response)
    # 各校 中 各班的 成績
    soc.class<-soc.all   %>%
      filter(缺考   == "0" & 補考   == "0" & 資源班   == "0")  %>%
      group_by(縣市,學校代碼,學校名稱,測驗科目代碼,年級,班級代碼)   %>%
      summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,15), na.rm = FALSE, .groups = 'drop_last')
    # 各校的 成績
    soc.school<-soc.all   %>%
      filter(缺考   == "0" & 補考   == "0" & 資源班   == "0")  %>%
      group_by(縣市,學校代碼,學校名稱,測驗科目代碼,年級) %>%
      summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,15), na.rm = FALSE, .groups = 'drop_last') 
    #  各縣市 的 成績
    soc.counties<-soc.all   %>%
      filter(缺考   == "0" & 補考   == "0" & 資源班   == "0")  %>% 
      group_by(縣市,測驗科目代碼,年級)     %>%
      summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,15), na.rm = FALSE, .groups = 'drop_last')
    #  全體 的 成績
    soc.total<-soc.all   %>%
      filter(缺考   == "0" & 補考   == "0" & 資源班   == "0")  %>% 
      group_by(測驗科目代碼,年級)     %>%
      summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,15), na.rm = FALSE, .groups = 'drop_last')
    ########################################
    ########################################
    #  匯出資料
    ########################################
    ########################################
    wb <- createWorkbook()
    
    #  addWorksheet(wb, sheetName = "原始資料")
    addWorksheet(wb, sheetName = "個人成績計算")
    addWorksheet(wb, sheetName = "各校各班通過率")
    addWorksheet(wb, sheetName = "各校通過率")
    addWorksheet(wb, sheetName = "縣市通過率")
    addWorksheet(wb, sheetName = "全體通過率")
    
    #  writeData(wb, sheet = "原始資料"       , x = df.s)
    writeData(wb, sheet = "個人成績計算"   , x = soc.person)
    writeData(wb, sheet = "各校各班通過率" , x = soc.class)
    writeData(wb, sheet = "各校通過率"     , x = soc.school)
    writeData(wb, sheet = "縣市通過率"     , x = soc.counties)
    writeData(wb, sheet = "全體通過率"     , x = soc.total)
    # writeData(wb, sheet = "缺考名單"       , x = df.miss)
    
    saveWorkbook(wb, paste( "C:/Users/user/Desktop/",counties.name,"/",counties.name,科目種類[[j,1]],科目種類[[j,2]],"(含班級).xlsx" ,  sep=""), overwrite = TRUE)
    # 個人成績計算
    # 成績匯出
    # write_excel_csv(soc.person,
    #                 file = paste( "C:/Users/user/Desktop/嘉義縣",科目種類[[j,1]],科目種類[[j,2]],".csv" ,  sep=""),
    #                 sep  = ",",
    #                 bom  = T )
    # 
    # 
    # 
    # fwrite(soc.person,
    #        file = paste( "C:/Users/user/Desktop/嘉義縣",科目種類[[j,1]],科目種類[[j,2]],".csv" ,  sep=""),
    #        sep  = ",",
    #        bom  = T )
    # 
    #  # 匯出檔案
    #  export(list(原始資料          = df.s,
    #              個人成績計算      = soc.person,
    #              各校各班通過率    = soc.class,
    #              各校通過率        = soc.school,
    #              縣市通過率        = soc.total,
    #              缺考名單          = df.miss), 
    #              paste( "C:/Users/user/Desktop/",counties.name,"/",counties.name,科目種類[[j,1]],科目種類[[j,2]],".xlsx" ,  sep=""))
    #  
    
    
    
  } #第二迴圈 各科目 與 年級
} #第一迴圈 縣市

)
