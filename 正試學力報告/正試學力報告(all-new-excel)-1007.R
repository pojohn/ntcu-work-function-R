library(tidyverse)
library(openxlsx)
library(readxl)
library(data.table)
library(ggrepel)
library(gridExtra)
library(grid)
# "C:/Users/user/Documents/台中教育大學/學力工作/20200827-學力表現與各區答對率/csv"
# "E:\\台中教育大學\\學力工作\\20200827-學力表現與各區答對率/csv"

# counties<-new.path %>% basename() %>% strsplit( "[成績,. ]") # 保留部分檔案名 用來存檔

# 合併 各檔案 製作 總檔
# tbl = lapply(new.path, function(x) fread(x,select=list(character=1:21,numeric = 22),encoding = "UTF-8" ) ) %>% bind_rows()
# fwrite(x = tbl,file = "C:/Users/user/Desktop/00all.csv",na = "")


# 前期設定
{
  old.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200827-學力表現與各區答對率/1007-最新程式分割檔案"
  new.path<-list.files(path = old.path,full.names = T)#指定資料夾內的資料夾清單
##############################################################################################################################
#  開始前 重置
df.all<-data.table()
# 讀入 總檔案   
df.all <- fread(new.path[1],na.strings = "",encoding = "UTF-8")
# 將 NA部分處理為 0
df.all[is.na(df.all)  == TRUE]  <- 0

# 抓取 資料的 縣市種類 並 排序
df.all$縣市 <- factor(df.all$縣市, levels=unique(df.all$縣市))
縣市種類    <- df.all[,.(N = .N),by = .(縣市)] %>% setorder(縣市)

科目排序    <- c("國語文","英語文","數學","自然")
df.all$科目 <- factor(df.all$科目, levels=unique(科目排序))
科目種類    <- df.all[,.(N = .N),by = .(科目)] %>% setorder(科目)


年級排序    <- c("一年級","二年級","三年級","四年級","五年級","六年級","七年級","八年級")
df.all$年級 <- factor(df.all$年級, levels=unique(年級排序))
幾年級      <- df.all[,.(N = .N),by = .(年級)]  %>% setorder(年級)

國小        <- c("一年級","二年級","三年級","四年級","五年級","六年級")
國中        <- c("七年級","八年級")


########## 根據各身分欄位 新增 身分 欄位 
身份 <- 0
身份[df.all$原住民 == 0 & df.all$新住民 == 0] <- "一般生"
身份[df.all$原住民 == 1 & df.all$新住民 == 0] <- "原住民" 
身份[df.all$原住民 == 0 & df.all$新住民 == 1] <- "新住民"
身份[df.all$原住民 == 1 & df.all$新住民 == 1] <- "原+新住民"
身份 <- factor(身份,levels= c("一般生","原住民","新住民","原+新住民")) 
身份 <- data.frame( 身份)
# table(身份) 
########## 將 身分 合併進 資料
df.all.id<- data.table(df.all[,1:20],身份,df.all[,21])


##########  輸出表格設定 1
hs1 <- createStyle(fontColour = "#000000", fgFill = "#ffffff",
                   halign = "center", valign = "center", textDecoration = "bold",
                   border = "TopBottomLeftRight")
# 表格字體設定
windowsFonts(A=windowsFont("標楷體"))
# gridExtra 表格(圖) 字體設定 # https://www.rdocumentation.org/packages/gridExtra/versions/2.3/topics/tableGrob
mytheme <- gridExtra::ttheme_default( base_family = "A" )

# 讀入 CFA 總檔案
old.CFA.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200918-學力報告製作/問卷資料"
new.CFA.path<-list.files(path = old.CFA.path,full.names = T)#指定資料夾內的資料夾清單
df.CFA <- read_excel(new.CFA.path[2] , sheet = 1 , col_names = TRUE , na = "" ) %>% data.table()
df.CFA.p <- df.CFA   %>% filter(`顯著性`   <=0.05 )
# 替換 CFA 資料 
{ df.CFA.p$`Pearson 相關`  <- round(df.CFA.p$`Pearson 相關`+0.0000000000000001,2)
  # 將 NA部分處理為 0
  df.CFA.p[is.na(df.CFA.p)  == TRUE]  <- 0
  # 替換 年級
  df.CFA.p[["年級"]] <- gsub("1$",  paste( "一年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("2$",  paste( "二年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("3$",  paste( "三年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("4$",  paste( "四年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("5$",  paste( "五年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("6$",  paste( "六年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("7$",  paste( "七年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("8$",  paste( "八年級",  sep=""),df.CFA.p[["年級"]])
  df.CFA.p[["年級"]] <- gsub("9$",  paste( "九年級",  sep=""),df.CFA.p[["年級"]])
  # 替換 科目
  df.CFA.p[["科目"]] <- gsub("CHI$" ,  paste( "國語文",  sep=""),df.CFA.p[["科目"]])
  df.CFA.p[["科目"]] <- gsub("ENG$" ,  paste( "英語文",  sep=""),df.CFA.p[["科目"]])
  df.CFA.p[["科目"]] <- gsub("MATH$",  paste( "數學"  ,  sep=""),df.CFA.p[["科目"]])
  df.CFA.p[["科目"]] <- gsub("SCI$" ,  paste( "自然"  ,  sep=""),df.CFA.p[["科目"]])
} 
# 抓取 資料的 縣市種類 並 排序
df.CFA.p$科目 <- factor(df.CFA.p$科目, levels=unique(科目排序))

df.CFA.p$年級 <- factor(df.CFA.p$年級, levels=unique(年級排序))
df.CFA.p$縣市 <- factor(df.CFA.p$縣市, levels=unique(df.all$縣市))

}

# 各縣市 # 1:nrow(縣市種類) # c(2,3,7,8,13,15)
for (a in 1:nrow(縣市種類)){ 
  #開啟 新的 Workbook
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = "家庭背景")
  addWorksheet(wb, sheetName = "各年段表現")  
  addWorksheet(wb, sheetName = "各卷人數與成績相關") 
  addWorksheet(wb, sheetName = "各卷學力表現") 
  addWorksheet(wb, sheetName = "問卷相關係數") 
  # 指定 縣市
  df.all.idc <- df.all.id[縣市 == 縣市種類[[a,1]] ]
  # 該縣市 的 年級種類  for b  for e  for g
  幾年級.idc    <- df.all.idc[,.(N = .N),by = .(年級)] %>% setorder(年級)
  # 該縣市 的 科目種類  for c  
  科目種類.idc    <- df.all.idc[,.(N = .N),by = .(科目)] %>% setorder(科目)
  # 該縣市 的 年級 * 科目 種類  for d
  年級科目種類.idc    <- df.all.idc[,.(N = .N),by = .(年級,科目)] %>% setorder(年級,科目)
  # 判斷 CFA 資料 有無該縣市
  if( nrow( df.CFA.p [縣市 == 縣市種類[[a,1]] ]  ) > 5  ){
    df.CFA.p.idc <- df.CFA.p [縣市 == 縣市種類[[a,1]] ]
    年級科目種類.CFA.idc    <- df.CFA.p.idc[,.(N = .N),by = .(年級,科目)] %>% setorder(年級,科目)
    年級種類.CFA.idc    <- df.CFA.p.idc[,.(N = .N),by = .(年級)] %>% setorder(年級)
  }
########################################
########################################
# 分頁: 家庭背景
########################################
########################################  
  # 1:nrow( 幾年級.idc )
  # 繪圖開關，僅測試表格時關閉
  trun.1 <- "1"
for(b in 1:nrow( 幾年級.idc ) ) { # 1:nrow( 幾年級.idc )
  # 指定 年級
  df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[b,1]] ]
 
  #1 人數
  df.Identity.1 <- df.all.idcy   %>%
    group_by(科目,身份) %>%
    summarise(人數 = n(), .groups = 'drop_last' )   %>% data.table()
  # 在科目尾部 加上 "-人數" 文字
  df.Identity.1$科目 <- gsub("$", "-人數",df.Identity.1[,科目])
  # 將 各地區 人數 長轉寬 表格 #spread(df.Identity.1, key = 科目  , value = 人數)
  df.Identity.1w <- pivot_wider(df.Identity.1,names_from = 科目, values_from = 人數)
  
  
  
  #2 平均數
  df.Identity.2 <- df.all.idcy    %>%
    group_by(科目,身份) %>%
    summarise( 總平均數 = round(mean(總平均)+0.0000000000000001,2) , .groups = 'drop_last' )   %>% data.table()
  # 在科目尾部 加上 "-平均數" 文字
  df.Identity.2$科目 <- gsub("$", "平均數",df.Identity.2[,科目])
  # 將 各地區 總平均數 長轉寬 表格 #spread(df.Identity.2, key = 科目  , value = 總平均數)
  df.Identity.2w <- pivot_wider(df.Identity.2,names_from = 科目, values_from = 總平均數)
  
  # 合併 人數 + 平均數
  df.Identity.ww <- data.table(df.Identity.1w[,-1] ,df.Identity.2w)
  # 改 欄位名稱 身份
  names(df.Identity.ww)[ncol(df.Identity.2w)] <- paste( 幾年級.idc[[b,1]], sep = "")
  #輸出 檔案 至  分頁: 家庭背景
  writeData(wb, sheet = "家庭背景" ,
            startCol = (6-ncol(df.Identity.2w)),
            startRow = (24*b-22),
            headerStyle = hs1 ,borders = "all", x = df.Identity.ww)
  
  
  
  
  

  if (trun.1 == "1"){
  #繪圖設定
  g1.text          = 4
  g1.axis.text.x   = 14
  g1.width         = 22
  g1.height        = 11
  g1.plot.title    = 18
  g1.axis.title.xy = 16
  g1.legend.title  = 22
  g1.legend.text   = 12
  g1.strip.text.x  = 12
  windowsFonts(A=windowsFont("標楷體"))
  thm <- function() theme(axis.text.x  = element_text(angle = 0, vjust = 0.5,size= g1.axis.text.x ,family = "A"),
                          axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                          strip.text.x = element_text(size = g1.strip.text.x ,family = "A"),
                          plot.title   = element_text(hjust = 0.5,size= g1.plot.title ,family = "A"),
                          axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g1.axis.title.xy,family = "A"),
                          axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g1.axis.title.xy,family = "A"),
                          legend.title = element_text(colour="black", size=g1.legend.title, face="bold",family = "A"),
                          legend.text  = element_text(colour="black", size = g1.legend.text,family = "A")) # Windows user
  # 長條圖
  p1 <- ggplot(data=df.Identity.2, aes(x=身份, y= 總平均數  ,fill=科目,label=總平均數 )) +
    geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
    geom_text(hjust= 0.5 , vjust = 1.6,color="white",
              position = position_dodge(0.9), size=g1.text)+
    labs(title = paste( 縣市種類[[a,1]],"不同家庭背景",幾年級.idc[[b,1]],"學生學力表現", sep = ""),
         y = "平\n均\n數",family = "A")+
    guides(    color = guide_colorbar(order = 0),
               fill = guide_legend(order = 1)  )+
    theme_minimal() + thm()
  
  ## 匯入圖片
  print(p1) # plot needs to be showing
  insertPlot(wb, sheet = "家庭背景",startCol = 11,startRow = (24*b-22), width = g1.width, height = g1.height, fileType = "png", units = "cm")
}
} 
  
########################################
########################################
# 分頁: 各年段表現
########################################
########################################  
  # 1:nrow( 科目種類.idc )
  # 繪圖開關，僅測試表格時關閉
  trun.2 <- "1"
for(c in 1:nrow( 科目種類.idc ))  {
  # 指定 科目
  df.all.idcs <- df.all.idc[科目 == 科目種類.idc[[c,1]] ]
  #1 年級
  df.year.1 <- df.all.idcs   %>%
    group_by(年級,地區) %>%
    summarise(人數 = n(), .groups = 'drop_last' )   %>% data.table()

  # 將 各地區 人數 長轉寬 表格 # spread(df.year.1, key = 年級  , value = 人數)
  df.year.1w <-pivot_wider(df.year.1,names_from = 年級, values_from = 人數)
  # 在年級尾部 加上 "-人數" 文字
  names(df.year.1w) <- gsub("年級$", "年級-人數",names(df.year.1w))
  
  #2 平均數
  df.year.2 <- df.all.idcs    %>%
    group_by(年級,地區) %>%
    summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) , .groups = 'drop_last' )  %>% data.table()
  

  # 將 各地區 平均數 長轉寬 表格 # spread(df.year.2, key = 年級  , value = 平均數)
  df.year.2w <-   pivot_wider(df.year.2,names_from = 年級, values_from = 平均數)
  # 在年級尾部 加上 "-人數" 文字
  names(df.year.2w) <- gsub("年級$", "年級平均數",names(df.year.2w))
  
  df.es.ms <- df.all.idcs   %>%
    group_by(地區) %>%
    summarise(總人數 = n(), .groups = 'drop_last' )
  # 合併 人數 + 平均數
  df.year.ww <- data.table(df.es.ms[,-1],df.year.1w[,-1] ,df.year.2w) %>% setorder(總人數 )
  # 移除 總人數欄位
  df.year.ww <-   df.year.ww[,-1]
  # 改 欄位名稱 身份
  names(df.year.ww)[ncol(df.year.2w)] <- paste( 科目種類.idc[[c,1]],"-地區", sep = "")

  # 確認 "地區" 數量是否 超過 預設範圍
  if (nrow(df.year.ww) <= 13 ){
    p2.sc      <- (26*c-24)
    g2.width   <- 26
    g2.height  <- 13
    
  }else{if(nrow(df.year.ww)<=20){
    p2.sc      <- (32*c-30)
    g2.width   <- 30
    g2.height  <- 15
  }else{if(nrow(df.year.ww) <= 30){
    p2.sc      <- (42*c-40)
    g2.width   <- 32
    g2.height  <- 15
  }else{
    p2.sc      <- (52*c-50)
    g2.width   <- 34
    g2.height  <- 15
  }            }          }
  #輸出 檔案 至  分頁: 家庭背景
  writeData(wb, sheet = "各年段表現" ,
            startCol = (1),
            startRow = p2.sc,
            headerStyle = hs1 ,borders = "all", x = df.year.ww)
  
  
  # 畫圖用 人數排序
  df.es.ms <- df.all.idcs   %>%
    group_by(地區) %>%
    summarise(人數 = n(), .groups = 'drop_last' )   %>% setorder(人數) %>% data.table()
  df.es.ms.order <- df.es.ms
  df.es.ms.order[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.es.ms.order[["地區"]] )
  # 國小 1~6 年級
  df.year.2.es <- filter(df.year.2,年級 %in% 國小)
  #  設定 變數 的排序順序
  df.year.2.es[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.year.2.es[["地區"]] )
  df.year.2.es[["地區"]] <- factor(df.year.2.es[["地區"]],levels = unique(df.es.ms.order[["地區"]]) )
  # 國中 7~8 年級
  df.year.2.ms <- filter(df.year.2,年級 %in% 國中)
  #  設定 變數 的排序順序
  df.year.2.ms[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.year.2.ms[["地區"]] )
  df.year.2.ms[["地區"]] <- factor(df.year.2.ms[["地區"]],levels = unique(df.es.ms.order[["地區"]]) )



  if (trun.2 == "1"){
    #繪圖設定
    g2.text          = 4
    g2.axis.text.x   = 14
    g2.plot.title    = 18
    g2.axis.title.xy = 16
    g2.legend.title  = 22
    g2.legend.text   = 12
    g2.strip.text.x  = 12
    x.angle = 0
    windowsFonts(A=windowsFont("標楷體"))
    thm2 <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g2.axis.text.x ,family = "A"),
                             axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                             strip.text.x = element_text(size = g2.strip.text.x ,family = "A"),
                             plot.title   = element_text(hjust = 0.5,size= g2.plot.title ,family = "A"),
                             axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g2.axis.title.xy,family = "A"),
                             axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g2.axis.title.xy,family = "A"),
                             legend.title = element_text(colour="black", size=g2.legend.title, face="bold",family = "A"),
                             legend.text  = element_text(colour="black", size = g2.legend.text,family = "A"),# Windows user
                             legend.position="bottom") 
    # 折線圖
    p2.es <- ggplot(data=df.year.2.es, aes(x=地區, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
      geom_line(size=1)+
      geom_point(size=1)+
     # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
      labs(title = paste(   縣市種類[[a,1]],"國小各區域",科目種類.idc[[c,1]],"學力表現", sep = ""),
           x = "區域參與人數由少➜多" ,
           y = "平\n均\n數",family = "A")+
      theme_minimal() + thm2()
    

    ## 匯入圖片
    # 國小 es
    print(p2.es) # plot needs to be showing
    insertPlot(wb, sheet = "各年段表現",
               startCol = 18,startRow = p2.sc ,
               width = g2.width, height = g2.height, fileType = "png", units = "cm")
    
    # 國中 ms
    if(nrow(df.year.2.ms)>1){
      ms.p2 <- ggplot(data=df.year.2.ms, aes(x=地區, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
        geom_line(size=1)+
        geom_point(size=1)+
        # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
        labs(title = paste(   縣市種類[[a,1]],"國中各區域",科目種類.idc[[c,1]],"學力表現", sep = ""),
             x = "區域參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        theme_minimal() + thm2()
      
      
      ## 匯入圖片
      print(ms.p2) # plot needs to be showing
      insertPlot(wb, sheet = "各年段表現",
                 startCol = (g2.width+10),startRow = p2.sc ,
                 width = g2.width, height = g2.height, fileType = "png", units = "cm")
    }
  }
  
}
  
########################################
########################################
# 分頁: 各卷人數與成績相關
########################################
########################################  
  # 1:nrow( 年級科目種類.idc )
for(d in 1:nrow( 年級科目種類.idc ) )  {
  #  開始前 重置
  df.single <- data.table()
  df.single.path <- paste( old.path,"/",年級科目種類.idc[[d,2]],年級科目種類.idc[[d,1]],".csv" ,  sep="")
  # 讀入 單 科目*年級 檔案   
  df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")

  # 指定 單檔案  的縣市
  df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
  # 排除 計算不需要的 欄位
  df.single.cs  <- data.table(df.single.c[,c("地區")],df.single.c[,c(-1:-20)])
  

  
  
  
  # 平均
  cor.sc.01 <- df.single.cs    %>%
    group_by(地區)       %>%
    summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()
  # 人數
  cor.sc.02 <- df.single.cs    %>%
    group_by(地區)  %>%
    summarise(人數 = n(), .groups = 'drop_last' )   %>% data.table()
  # 標準差
  cor.sc.03 <- df.single.cs    %>%
    group_by(地區)       %>%
    summarise_if(is.numeric,~round(sd(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()
  
  # 合併 人數 * 平均
  cor.sc.04 <- full_join(cor.sc.02,cor.sc.01,by = "地區") %>% data.table()
  # 合併 人數 * 標準差
  cor.sc.05 <- full_join(cor.sc.02,cor.sc.03,by = "地區") %>% data.table()
  
  # 相關
  #1 平均
  cor.t1 <- data.table(Pearson相關 = "各區人數--答對率相關")
  cor.c1 <- cor(x = cor.sc.04[,人數] , y = cor.sc.04[,c(3:ncol(cor.sc.04)),with = FALSE], method = 'pearson') %>% round(digits = 2) %>% data.table()
  #2 標準差
  cor.t2 <- data.table(Pearson相關 = "各區人數--標準差相關")
  cor.c2 <- cor(x = cor.sc.05[,人數] , y = cor.sc.05[,c(3:ncol(cor.sc.04)),with = FALSE], method = 'pearson') %>% round(digits = 2)  %>% data.table()
  
  cor.ct1 <- data.table(cor.t1,cor.c1) 
  cor.ct2 <- data.table(cor.t2,cor.c2) 
  
  cor.ct3 <- bind_rows(cor.ct1,cor.ct2)
  # 在年級尾部 加上 "-人數" 文字
  names(cor.ct3) <- gsub("Pearson相關$",
                         paste( 年級科目種類.idc[[d,2]],年級科目種類.idc[[d,1]],"-Pearson相關",  sep=""),names(cor.ct3))
  #測試 表格寬度自適應
 # width_vec <- apply(cor.ct3, 2, function(x) max(nchar(as.character(x)) + 2, na.rm = TRUE))
  #輸出 檔案 至  分頁: 各卷人數與成績相關
  writeData(wb, sheet = "各卷人數與成績相關" ,
            startCol = (1),
            startRow = (6*d-5),
            headerStyle = hs1 ,borders = "all", x = cor.ct3)
}
  
########################################
########################################
# 分頁: 各卷學力表現
########################################
########################################    
  # 1:nrow( 幾年級.idc )
  # 1:nrow( 科目種類.idcs )
  # 繪圖開關，僅測試表格時關閉
  trun.3 <- "1"
for(e in 1:nrow( 幾年級.idc ) ) {
  # 指定 年級的 科目種類
  df.single.idcy <- df.all.idc[年級 == 幾年級.idc[[e,1]] ]
  科目種類.idcs    <- df.single.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
  
  for(f in 1:nrow( 科目種類.idcs ) ){
  #  開始前 重置
  df.single <- data.table() 
  df.single.path <-paste( old.path,"/",科目種類.idcs[[f,1]],幾年級.idc[[e,1]],".csv" ,  sep="")
  # 讀入 單 科目*年級 檔案   
  df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
  # 指定 單檔案  的縣市
  df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
  
  
  ########## 根據總分排名 新增 高、低分組 欄位
  # 先取得 高、低分組線(27%)
  low  <- quantile(df.single.c$總平均, probs=0.27)
  high <- quantile(df.single.c$總平均, probs=0.73)
  高低分組 <- 0
  高低分組[df.single.c$總平均 <= low] <- "低分組"
  高低分組[df.single.c$總平均 >  low  &  df.single.c$總平均 <high] <- "中間組" 
  高低分組[df.single.c$總平均 >= high] <- "高分組"
  高低分組 <- factor(高低分組, levels=c("高分組","中間組","低分組")) 
  高低分組 <-data.table(高低分組) 
  table(高低分組) 
  ########## 將 高、低分組 合併進 資料
  df.single.hl<- data.table(高低分組,df.single.c[,21:ncol(df.single.c)])

  # 平均
  hl.sc.01 <- df.single.hl    %>%
    summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()
  hl.sc.t1 <- data.table(計算類別 = "平均")
  hl.sc.ct1 <- data.table(hl.sc.t1,hl.sc.01)
  # 標準差
  hl.sc.02 <- df.single.hl    %>%
    summarise_if(is.numeric,~round(sd(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()  
  hl.sc.t2 <- data.table(計算類別 = "標準差")
  hl.sc.ct2 <- data.table(hl.sc.t2,hl.sc.02)
  # 高低分組
  df.single.h <- df.single.hl[高低分組 == "高分組"] %>%
    summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()
  df.single.l <- df.single.hl[高低分組 == "低分組"] %>%
    summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()
  hl.sc.03 <- df.single.h-df.single.l
  hl.sc.t3 <- data.table(計算類別 = "高分組-低分組")
  hl.sc.ct3 <- data.table(hl.sc.t3,hl.sc.03)
  # 合併表格 並 寬轉長
  hl.sc.ct4 <- bind_rows(hl.sc.ct1,hl.sc.ct2,hl.sc.ct3) 
  hl.sc.ct4$計算類別 <- factor(hl.sc.ct4$計算類別, levels=c("平均","標準差","高分組-低分組")) 
  # 將 各向度 通過率 寬轉長 表格
  sub.names<-paste(科目種類.idcs[[f,1]],幾年級.idc[[e,1]],"向度名稱" ,  sep="")
  hl.sc.ctp<-pivot_longer(hl.sc.ct4, cols = - 計算類別, names_to = sub.names, values_to = "平均數"  )
  hl.sc.ct5<-pivot_wider(hl.sc.ctp,names_from = 計算類別, values_from = 平均數)
  #  設定 變數 的排序順序
  hl.sc.ctp[[sub.names]] <- factor(hl.sc.ctp[[sub.names]],levels = unique(hl.sc.ct5[[sub.names]]) )

  #輸出 檔案 至  分頁: 各卷人數與成績相關
  writeData(wb, sheet = "各卷學力表現" ,
            startCol = (10*f-9),
            startRow = (43*e-42),
            headerStyle = hs1 ,borders = "all", x = hl.sc.ct5)
  
  if (trun.3 == "1"){
    #繪圖設定
    nl =unique(hl.sc.ctp[[sub.names]]) %>% length() 
    g3.text          = if(nl<=6 ){g3.text = 4 }else{g3.text = 2}
    g3.axis.text.x   = if(nl<=6 ){g3.axis.text.x = 14}else{g3.axis.text.x = 8}
    g3.width         = 22
    g3.height        = 11
    g3.plot.title    = 18
    g3.axis.title.xy = 16
    g3.legend.title  = 22
    g3.legend.text   = 12
    g3.strip.text.x  = 12
    windowsFonts(A=windowsFont("標楷體"))
    thm3 <- function() theme(axis.text.x  = element_text(angle = 20, vjust = 0.5,size= g3.axis.text.x ,family = "A"),
                            axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                            strip.text.x = element_text(size = g3.strip.text.x ,family = "A"),
                            plot.title   = element_text(hjust = 0.5,size= g3.plot.title ,family = "A"),
                            axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g3.axis.title.xy,family = "A"),
                            axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g3.axis.title.xy,family = "A"),
                            legend.title = element_text(colour="black", size=g3.legend.title, face="bold",family = "A"),
                            legend.text  = element_text(colour="black", size = g3.legend.text,family = "A")) # Windows user
    # 長條圖
    # geom_text_repel(hjust= 0.5 , vjust = 1.6,color="white",
    #                 position = position_dodge(0.9), size=g1.text)
    p3 <- ggplot(data=hl.sc.ctp, aes_string(x=`sub.names`, y= "平均數"  ,fill="計算類別",label="平均數" )) +
      geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
      geom_text(hjust= 0.5 , vjust = 1.6,color="white",
                position = position_dodge(0.9), size=g3.text)+
      scale_fill_manual(values=c("#2b47d2","#71ccfe","#fe0733" ))+
      labs(title = paste( 縣市種類[[a,1]],科目種類.idcs[[f,1]],幾年級.idc[[e,1]],"學力表現", sep = ""),
           x = "向度名稱",
           y = "平\n均\n數",family = "A")+
      guides(    color = guide_colorbar(order = 0),
                 fill = guide_legend(order = 1)  )+
      theme_minimal() + thm3()
    
    ## 匯入圖片
    print(p3) # plot needs to be showing
    insertPlot(wb, sheet = "各卷學力表現",
               startCol = (10*f-9),
               startRow = (43*e-27),
               width = g3.width, height = g3.height, fileType = "png", units = "cm")
  }
  
}}

########################################
########################################
# 分頁: 問卷相關係數
########################################
########################################  
  
if( nrow( df.CFA.p [縣市 == 縣市種類[[a,1]] ]  ) > 5  ){
    # CFA 問卷 與 成績 相關 (長條圖)  
  for(i in 1:nrow(年級種類.CFA.idc) )  {
    df.CFA.p.idcys <- df.CFA.p.idc[年級 == 年級種類.CFA.idc[[i,1]]  ] %>% setorder(-`Pearson 相關`)
    
    df.CFA.p.idcys$`向度`<- factor(df.CFA.p.idcys$`向度`, levels=unique(df.CFA.p.idcys$`向度`))
    df.CFA.p.idcys <- df.CFA.p.idcys[,1:6]
    
    df.CFA.p.idcys.L<- df.CFA.p.idcys   %>%
      group_by(科目,`向度`) %>%
      summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')   %>% data.table()
    # 將 各向度 長轉寬 表格 
    df.CFA.p.idcys.W <- pivot_wider(df.CFA.p.idcys.L,names_from = 科目, values_from = `Pearson 相關`)
    names(df.CFA.p.idcys.W)[1] <- paste( 年級種類.CFA.idc[[i,1]],"相關係數",  sep="")


    #輸出 檔案 至  分頁: 問卷相關係數
    writeData(wb, sheet = "問卷相關係數" ,
              startCol = 2,
              startRow = (i*15-13),
              headerStyle = hs1 ,borders = "all", x = df.CFA.p.idcys.W)
      
    }
  }
  
  
    
########################################
########################################
# 分頁: N年級整體學力表現
########################################
########################################   
  # 1:nrow( 幾年級.idc )
  # 1:nrow( 科目種類.idcs )
  # 繪圖開關，僅測試表格時關閉
  trun.4 <- "1"
  for(g in 1:nrow( 幾年級.idc ) ){
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[g,1]] ]
    # 指定 科目 for h
    科目種類.idcs    <- df.all.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
    # 建立 N年級整體學力表現 分頁   
    n.page = paste( 幾年級.idc[[g,1]],"整體學力表現", sep="")
    addWorksheet(wb, sheetName = n.page ) 
    
    
    #人數表
    df.all.person.l <- df.all.idcy    %>%
      group_by(科目,地區) %>%
      summarise(人數 = n() , .groups = 'drop_last' )  %>% setorder(人數) %>% data.table()
    # 將欄位 科目 更名為 科目人數
    names(df.all.person.l) <- c("科目-人數","地區" ,"人數")
    # 將 各地區 人數 長轉寬 表格 ( 要用 )
    df.all.person.w <-pivot_wider(df.all.person.l,names_from = 地區, values_from = 人數)
    
    # 折線圖資料 ( 要用 )
    g4.01 <- df.all.idcy    %>%
      group_by(地區,科目) %>%
      summarise(人數 = n(), 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                  標準差 = round(sd(總平均)+0.0000000000000001,2) , .groups = 'drop_last' ) %>% setorder(人數)
    #篩選  地區   科目   平均數
    g4.02 <- g4.01[,c(1:2,4)] 
    # 在科目尾部 加上 "平均數" 文字
    g4.02[["科目"]] <- gsub("$", paste( "平均數",  sep=""),g4.02[["科目"]])
    #折線圖 旁 資料表 長轉寬 表格 ( 要用 )
    g4.03 <- pivot_wider(g4.02,names_from = 科目, values_from = 平均數)
    
    # 確認 "地區" 數量是否 超過 預設範圍
    if (nrow(g4.03) <= 13 ){
      p0.sc      <- 20
      g4.width   <- 26
      g4.height  <- 13
    }else{if(nrow(g4.03)<=20){
      p0.sc      <- 28
      g4.width   <- 40
      g4.height  <- 15
    }else{if(nrow(g4.03) <= 30){
      p0.sc      <- 36
      g4.width   <- 55
      g4.height  <- 17
    }else{
      p0.sc      <- 45
      g4.width   <- 65
      g4.height  <- 19
    }            }          }
    
    
    #輸出 折線圖 旁 資料表 至  分頁: N 年級整體學力表現
    writeData(wb, sheet = n.page ,
              startCol = (1),
              startRow = p0.sc,
              headerStyle = hs1 ,borders = "all", x = g4.03)
    #輸出 人數表 至  分頁: N 年級整體學力表現
    writeData(wb, sheet = n.page ,
              startCol = (24),
              startRow = (3),
              headerStyle = hs1 ,borders = "all", x = df.all.person.w)
    if(trun.4 == "1"){
      g4.text          = 4
      g4.axis.text.x   = 14
      g4.plot.title    = 18
      g4.axis.title.xy = 16
      g4.legend.title  = 22
      g4.legend.text   = 12
      g4.strip.text.x  = 12
      windowsFonts(A=windowsFont("標楷體"))
      thm4 <- function() theme(axis.text.x  = element_text(angle = 0, vjust = 0.5,size= g4.axis.text.x ,family = "A"),
                               axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                               strip.text.x = element_text(size = g4.strip.text.x ,family = "A"),
                               plot.title   = element_text(hjust = 0.5,size= g4.plot.title ,family = "A"),
                               axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                               axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                               legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                               legend.text  = element_text(colour="black", size = g4.legend.text,family = "A"),# Windows user
                               legend.position="bottom") 
      # 修改 地區 字排列
      
      g4.01[["地區"]] <- gsub("*",  paste( "\n",  sep=""),g4.01[["地區"]] )
      g4.01[["地區"]] <- factor(g4.01[["地區"]],levels = unique(g4.01[["地區"]]) )
      p4<-ggplot(g4.01,aes(x=地區,y=平均數,group=科目,color=科目,label=平均數))+
        geom_line(size=1)+
        geom_point(size=1)+
        # geom_text_repel(size=4,color="black")+
        labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[g,1]],  "各區域整體學力表現", sep = ""),
             x = "區域參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        theme_minimal() + thm4()
      
      #vp <- viewport(x = 0.5, y = 0.5, width = 5, height = 5)
      tbl4   <- gridExtra::tableGrob(df.all.person.w, rows=NULL,theme = mytheme)
      
      grid.arrange(p4, tbl4, 
                   nrow = 2,
                   as.table = TRUE ,
                   heights = c(10,5 ) # ,widths =25
      )
      insertPlot(wb, sheet = n.page, startRow = p0.sc, startCol = 7, width = g4.width, height = g4.height, fileType = "png", units = "cm")    
      
    }
    
    for(h in 1:nrow( 科目種類.idcs )){
      #  開始前 重置
      df.single <- data.table() 
      df.single.path <-paste( old.path,"/",科目種類.idcs[[h,1]],幾年級.idc[[g,1]],".csv" ,  sep="")
      # 讀入 單 科目*年級 檔案   
      df.single   <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
      # 指定 單檔案  的縣市
      df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
      
      
      
      # 第一部分:
      {   soc.s.11 <- df.single.c    %>%
          group_by(地區)  %>%
          summarise(人數 = n(), 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                      標準差 = round(sd(總平均)+0.0000000000000001,2) , .groups = 'drop_last' ) %>% setorder(人數)
        
        
        # 在地區前面 加上 "該科目" 文字
        names(soc.s.11) <- gsub("地區$",
                                paste( 科目種類.idcs[[h,1]],"-地區",  sep=""),names(soc.s.11))
        #輸出 科目統計表 至  分頁: N 年級整體學力表現
        writeData(wb, sheet = n.page ,
                  startCol = (h*5-4),
                  startRow = (1),
                  headerStyle = hs1 ,borders = "all", x = soc.s.11)
        
      }
      
      #第二部分:
      {# 指定 "地區" 跟 各向度
        df.single.c2 <- df.single.c[,c(2,5,21:ncol(df.single.c)),with = FALSE]
        # 各向度平均數
        soc.s.21.m <- df.single.c2    %>%
          group_by(地區)  %>%
          summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last') 
        # 人數
        soc.s.21.p <- df.single.c2    %>%
          group_by(縣市,地區)  %>%
          summarise(人數 = n(), .groups = 'drop_last' )   %>%
          setorder(人數)
        # 合併 人數 * 各向度平均 (原始)
        soc.s.21 <- full_join(soc.s.21.p,soc.s.21.m,by = "地區") %>% data.table()
        
        
        #  表格用
        #1. 總答對率(低到高)排序
        soc.t.order  <- setorder(soc.s.21,總平均)
        soc.t.order$地區 <- factor(soc.t.order$地區, levels=unique(soc.t.order$地區))
        soc.t2.order <- soc.t.order[,c(-1,-3:-4)]

        
        #2. 區域人數(少到多)排序
        soc.p.order<-setorder(soc.s.21,人數)
        soc.p.order$地區 <- factor(soc.p.order$地區, levels=unique(soc.p.order$地區))
        soc.p2.order <- soc.p.order[,c(-1,-3:-4)]

        
        # 畫圖用 地區 * 人數表
        #1. 總答對率(低到高)排序
        soc.t3.order <- soc.t.order[,c(1:3)] %>% pivot_wider(names_from = 地區, values_from = 人數)
        #2. 區域人數(少到多)排序
        soc.p3.order <- soc.p.order[,c(1:3)] %>% pivot_wider(names_from = 地區, values_from = 人數)
        
        
        # 在 "縣市" 後面加上 "-人數"
        names(soc.t3.order) <- gsub("縣市$",  paste( 科目種類.idcs[[h,1]],"-縣市人數",  sep=""),names(soc.t3.order))
        names(soc.p3.order) <- gsub("縣市$",  paste( 科目種類.idcs[[h,1]],"-縣市人數",  sep=""),names(soc.p3.order))
        # 將 各向度 通過率 寬轉長 表格 (畫圖用)
        #1. 總答對率(低到高)排序
        soc.t.22<-pivot_longer(soc.t2.order, cols = -地區, names_to = "向度", values_to = "平均數"  )
        #2. 區域人數(少到多)排序
        soc.p.22<-pivot_longer(soc.p2.order, cols = -地區, names_to = "向度", values_to = "平均數"  )
        
        
        # 在地區前面 加上 "該科目" 文字
        #1. 總答對率(低到高)排序
        names(soc.t2.order) <- gsub("地區$",  paste( 科目種類.idcs[[h,1]],"-地區",  sep=""),names(soc.t2.order))
        #2. 區域人數(少到多)排序
        names(soc.p2.order) <- gsub("地區$",  paste( 科目種類.idcs[[h,1]],"-地區",  sep=""),names(soc.p2.order))

        
        # 確認 "地區" 數量是否 超過 預設範圍
        if (nrow(g4.03) <= 13 ){
          p4.sc      <- (h*60-10)
          p5.sc      <- (h*60+20)
          g4.width   <- 26
          g4.height  <- 13
        }else{if(nrow(g4.03)<=20){
          p4.sc      <- (h*80-10)
          p5.sc      <- (h*80+30)
          g4.width   <- 40
          g4.height  <- 15
        }else{if(nrow(g4.03) <= 30){
          p4.sc      <- (h*90-10)
          p5.sc      <- (h*90+35)
          g4.width   <- 55
          g4.height  <- 17
        }else{
          p4.sc      <- (h*120-20)
          p5.sc      <- (h*120+40)
          g4.width   <- 65
          g4.height  <- 19
        }            }          }
        #輸出1. 總答對率(低到高)排序 至  分頁: N 年級整體學力表現
        writeData(wb, sheet = n.page ,
                  startCol = (1),
                  startRow = p4.sc,
                  headerStyle = hs1 ,borders = "all", x = soc.t2.order)
        #輸出2. 區域人數(少到多)排序 至  分頁: N 年級整體學力表現
        writeData(wb, sheet = n.page ,
                  startCol = (1),
                  startRow = p5.sc,
                  headerStyle = hs1 ,borders = "all", x = soc.p2.order)
      }
      if(trun.4 == "1"){
        n.max = max(nchar(as.character(unique(soc.t.22[["向度"]]))) , na.rm = TRUE)
        nl    = unique(soc.t.22[["向度"]]) %>%length()
        
        g4.text          = 4
        g4.axis.text.x   = 14
        g4.plot.title    = 18
        g4.axis.title.xy = 16
        if(n.max<6){ g4.legend.title = 18 } else {g4.legend.title = 14}
        if(n.max<6){ g4.legend.text = 12 } else {g4.legend.text = 9}
        g4.strip.text.x  = 12
        
        if(nl<=6 & n.max<=5){gnrow = 1 } else{ if (nl<=10 & n.max<=6){gnrow = 2 } else{gnrow = 3}}
        
        thm4 <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g4.axis.text.x ,family = "A"),
                                 axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                 strip.text.x = element_text(size = g4.strip.text.x ,family = "A"),
                                 plot.title   = element_text(hjust = 0.5,size= g4.plot.title ,family = "A"),
                                 axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                                 axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                                 legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                                 legend.text  = element_text(colour="black", size = g4.legend.text,family = "A"),# Windows user
                                 legend.position="bottom"    ) 
        #1. 總答對率(低到高)排序
        soc.t.22[["地區"]] <- gsub("*",  paste( "\n",  sep=""),soc.t.22[["地區"]] )
        soc.t.22[["地區"]] <- factor(soc.t.22[["地區"]],levels = unique(soc.t.22[["地區"]]) )
        p5<-ggplot(soc.t.22,aes(x=地區,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[g,1]],科目種類.idcs[[h,1]],  "各區域學力表現", sep = ""),
               x = paste( 科目種類.idcs[[h,1]],  "區域平均答對率由低➜高", sep = ""),
               y = "平\n均\n數",family = "A")+
          guides( col   = guide_legend  (nrow  = gnrow  )  )+
          theme_minimal() + thm4()
        
        
        tbl5   <- gridExtra::tableGrob(soc.t3.order, rows=NULL,theme = mytheme)
        
        grid.arrange(p5, tbl5, 
                     nrow = 2,
                     as.table = TRUE ,
                     heights = c(10,2 )      )
        insertPlot(wb, sheet = n.page,
                   startRow = p4.sc,
                   startCol = 18, width = g4.width, height = g4.height, fileType = "png", units = "cm")
        
        #2. 區域人數(少到多)排序
        soc.p.22[["地區"]] <- gsub("*",  paste( "\n",  sep=""),soc.p.22[["地區"]] )
        soc.p.22[["地區"]] <- factor(soc.p.22[["地區"]],levels = unique(soc.p.22[["地區"]]) )
        p6<-ggplot(soc.p.22,aes(x=地區,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[g,1]],科目種類.idcs[[h,1]],  "各區域學力表現", sep = ""),
               x = "區域參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          guides( col   = guide_legend  (nrow  = gnrow  )  )+
          theme_minimal() + thm4()
        
        
        tbl6   <- gridExtra::tableGrob(soc.p3.order, rows=NULL,theme = mytheme)
        
        grid.arrange(p6, tbl6, 
                     nrow = 2,
                     as.table = TRUE ,
                     heights = c(10,2 )      )
        insertPlot(wb, sheet = n.page,
                   startRow = p5.sc,
                   startCol = 18, width = g4.width, height = g4.height, fileType = "png", units = "cm")
        
        
        
        
      }
      
    }
  }
  # 輸出檔案 各年段表現
  setColWidths(wb, sheet = "家庭背景"            , cols = 1:10, widths = "auto")
  setColWidths(wb, sheet = "各年段表現"          , cols = 1:18, widths = "auto")
  setColWidths(wb, sheet = "各卷人數與成績相關"  , cols = 1   , widths =  30) #只拉寬 標題的寬度，因為有向度字數差距過多
  setColWidths(wb, sheet = "各卷學力表現"        , cols = 1:35, widths =  "auto") 
  setColWidths(wb, sheet = "問卷相關係數"        , cols = 1:10, widths = "auto")
  saveWorkbook(wb, paste( "C:/Users/user/Desktop/測試/",縣市種類[[a,1]],"-1020.xlsx" ,  sep=""), overwrite = TRUE)
}

