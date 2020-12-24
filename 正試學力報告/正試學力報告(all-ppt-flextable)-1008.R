# ppt 設定的 套件網站
# https://davidgohel.github.io/flextable/index.html
# https://davidgohel.github.io/officer/index.html
# https://cran.r-project.org/web/packages/customLayout/vignettes/layouts-for-officer-power-point-document.html
library(tidyverse)
library(openxlsx)
library(readxl)
library(data.table)
library(ggrepel)
library(gridExtra)
library(grid)
library(officer)
library(customLayout)
library(flextable)

# counties<-new.path %>% basename() %>% strsplit( "[成績,. ]") # 保留部分檔案名 用來存檔
# 合併 各檔案 製作 總檔
# tbl = lapply(new.path, function(x) fread(x,select=list(character=1:21,numeric = 22),encoding = "UTF-8" ) ) %>% bind_rows()
# fwrite(x = tbl,file = "C:/Users/user/Desktop/00all.csv",na = "")


# 前期設定
{# "C:/Users/user/Documents/台中教育大學/學力工作/20200827-學力表現與各區答對率/csv"
  # "E:\\台中教育大學\\學力工作\\20200827-學力表現與各區答對率/csv"
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
  
  # 讀入 CFA 總檔案
  old.CFA.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200918-學力報告製作/問卷資料"
  new.CFA.path<-list.files(path = old.CFA.path,full.names = T)#指定資料夾內的資料夾清單
  df.CFA <- read_excel(new.CFA.path[2] , sheet = 1 , col_names = TRUE , na = "" ) %>% data.table()
  df.CFA.p <- df.CFA   %>% filter(`顯著性`   <=0.05 )
  # 替換 CFA 資料 
  { 
    df.CFA.p$`Pearson 相關`  <- round(df.CFA.p$`Pearson 相關` +0.0000000000000001,2)
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
  
  trun.1 <- "1" # 縣市 "各區域" 科目 "學力表現" (折線圖)
  trun.2 <- "1" # 縣市 幾年級  科目  "各區域學力表現" (折線圖)
  trun.3 <- "1" # 縣市 年級 科目 "學力表現" (長條圖)
  trun.4 <- "1" # 縣市 不同家庭背景   年級 學生學力表現
  trun.5 <- "1" # 縣市 年級 科目 相關表
}

# 各縣市 # 1:nrow(縣市種類) # c(2,3,7,8,13,15)
for (a in 1:nrow(縣市種類) ){
  # 指定 縣市
  df.all.idc   <- df.all.id[縣市 == 縣市種類[[a,1]] ]
  
  
  # 判斷 CFA 資料 有無該縣市
  if( nrow( df.CFA.p [縣市 == 縣市種類[[a,1]] ]  ) > 5  ){
    df.CFA.p.idc <- df.CFA.p [縣市 == 縣市種類[[a,1]] ]
    年級科目種類.CFA.idc    <- df.CFA.p.idc[,.(N = .N),by = .(年級,科目)] %>% setorder(年級,科目)
  }
  
  # 該縣市 的 年級種類  for b  for e  for g
  幾年級.idc        <- df.all.idc[,.(N = .N),by = .(年級)] %>% setorder(年級)
  # 幾年級.CFA.idc    <- df.CFA.p.idc[,.(N = .N),by = .(年級)] %>% setorder(年級)
  # 該縣市 的 科目種類  for c  
  科目種類.idc        <- df.all.idc[,.(N = .N),by = .(科目)] %>% setorder(科目)
  # 科目種類.CFA.idc    <- df.CFA.p.idc[,.(N = .N),by = .(科目)] %>% setorder(科目)
  # 該縣市 的 年級 * 科目 種類  for d
  年級科目種類.idc    <- df.all.idc[,.(N = .N),by = .(年級,科目)] %>% setorder(年級,科目)
  
  # 建立繪圖存檔用 資料夾
  # dir.create(paste("C:/Users/user/Desktop/測試/",縣市種類[[a,1]],sep = ""))
  #建立 ppt
  my_ppt <- read_pptx()
  
  
  
  
  # 縣市  科目 各區域 "學力表現" (折線圖)
  for(b in 1:nrow( 科目種類.idc )) {
    # 縣市 "各區域" 科目 "學力表現"
    if (trun.1 == "1"){
      # 指定 科目
      df.all.idcs <- df.all.idc[科目 == 科目種類.idc[[b,1]] ]
      
      df.all.idcs[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.all.idcs[["地區"]])
      #2 平均數
      df.year <- df.all.idcs    %>%
        group_by(年級,地區) %>%
        summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                  人數 = n(), .groups = 'drop_last' )   %>% setorder(人數)    %>% data.table()
      
      df.year$地區 <- factor(df.year$地區, levels=unique(df.year$地區))
      
      nl = length(unique(df.year$地區))
      
      # 國小 1~6 年級
      df.year.es <- filter(df.year,年級 %in% 國小)
      # 國中 7~8 年級
      df.year.ms <- filter(df.year,年級 %in% 國中)
      
      # 確認 "地區" 數量是否 超過 預設範圍
      if ( nl <= 13 ){
        p2.sc      <- (26*b-24)
        g2.width   <- 26
        g2.height  <- 13
        
      }else{if(nl <= 20){
        p2.sc      <- (32*b-30)
        g2.width   <- 30
        g2.height  <- 15
      }else{if( nl <= 30){
        p2.sc      <- (42*b-40)
        g2.width   <- 38
        g2.height  <- 18
      }else{
        p2.sc      <- (52*b-50)
        g2.width   <- 40
        g2.height  <- 20
      }            }          }
      
      #繪圖設定
      { g2.text          = 4
        if(nl<25){ g.axis.text.x = 14 } else {g.axis.text.x = 12}
        g2.plot.title    = 18
        g2.axis.title.xy = 16
        g2.legend.title  = 16
        g2.legend.text   = 12
        g2.strip.text.x  = 12
        x.angle = 0
        thm2 <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                                 axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                 strip.text.x = element_text(size = g2.strip.text.x ,family = "A"),
                                 plot.title   = element_text(hjust = 0.5,size= g2.plot.title ,family = "A"),
                                 axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g2.axis.title.xy,family = "A"),
                                 axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g2.axis.title.xy,family = "A"),
                                 legend.title = element_text(colour="black", size=g2.legend.title, face="bold",family = "A"),
                                 legend.text  = element_text(colour="black", size = g2.legend.text,family = "A"),# Windows user
                                 legend.position="bottom") }
      
      # 折線圖
      
      # 國小 es
      es.p2 <- ggplot(data=df.year.es, aes(x=地區, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
        geom_line(size=1)+
        geom_point(size=1)+
        # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
        labs(title = paste(   縣市種類[[a,1]],"國小各區域",科目種類.idc[[b,1]],"學力表現", sep = ""),
             x = "區域參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        theme_minimal() + thm2()
      
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = es.p2,
                        location = ph_location_fullsize() )
      # 國中 ms
      if(nrow(df.year.ms)>1){
        ms.p2 <- ggplot(data=df.year.ms, aes(x=地區, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
          labs(title = paste(   縣市種類[[a,1]],"國中各區域",科目種類.idc[[b,1]],"學力表現", sep = ""),
               x = "區域參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          theme_minimal() + thm2()
        
        
        ## 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = ms.p2,
                          location = ph_location_fullsize() )
      }
      
      
    }
  }
  
  # 縣市  科目 各校 "學力表現" (折線圖)
  for(b in 1:nrow( 科目種類.idc )) {
    # 縣市 "各區域" 科目 "學力表現"
    if (trun.1 == "1"){
      # 指定 科目
      df.all.idcs <- df.all.idc[科目 == 科目種類.idc[[b,1]] ]
      
      
      #2 平均數
      df.year <- df.all.idcs    %>%
        group_by(年級,學校ID) %>%
        summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                  人數 = n(), .groups = 'drop_last' )   %>% setorder(人數)    %>% data.table()
      
      df.year$學校ID <- factor(df.year$學校ID, levels=unique(df.year$學校ID))
      
      # 國小 1~6 年級
      df.year.es <- filter(df.year,年級 %in% 國小)
      # 國中 7~8 年級
      df.year.ms <- filter(df.year,年級 %in% 國中)
      # 確認 "學校數" 數量是否 超過 預設範圍
      if (nrow(df.year) <= 13 ){
        g2.width   <- 26
        g2.height  <- 13
        
      }else{if(nrow(df.year)<=20){
        g2.width   <- 30
        g2.height  <- 15
      }else{if(nrow(df.year) <= 30){
        
        g2.width   <- 38
        g2.height  <- 18
      }else{
        g2.width   <- 40
        g2.height  <- 20
      }            }          }
      
      #繪圖設定
      {
        g2.text          = 4
        if(nrow(df.year)<25){ g.axis.text.x = 14 } else {g.axis.text.x = 12}
        g2.plot.title    = 18
        g2.axis.title.xy = 16
        g2.legend.title  = 16
        g2.legend.text   = 12
        g2.strip.text.x  = 12
        x.angle = 0
        thm2 <- function() theme(axis.text.x  = element_text(colour="white",angle = x.angle, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                                 axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                 strip.text.x = element_text(size = g2.strip.text.x ,family = "A"),
                                 plot.title   = element_text(hjust = 0.5,size= g2.plot.title ,family = "A"),
                                 axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g2.axis.title.xy,family = "A"),
                                 axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g2.axis.title.xy,family = "A"),
                                 legend.title = element_text(colour="black", size=g2.legend.title, face="bold",family = "A"),
                                 legend.text  = element_text(colour="black", size = g2.legend.text,family = "A"),# Windows user
                                 legend.position="bottom") 
      }
      
      # 折線圖
      # 國小 es
      es.p2 <- ggplot(data=df.year.es, aes(x=學校ID, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
        geom_line(size=0.5)+
        geom_point(size=0.5)+
        # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
        labs(title = paste(   縣市種類[[a,1]],"國小各校",科目種類.idc[[b,1]],"學力表現", sep = ""),
             x = "學校參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        theme_minimal() + thm2()
      
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = es.p2,
                        location = ph_location_fullsize() )
      
      # 國中 ms
      if(nrow(df.year.ms)>1){
        ms.p2 <- ggplot(data=df.year.ms, aes(x=學校ID, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
          geom_line(size=0.5)+
          geom_point(size=0.5)+
          # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
          labs(title = paste(   縣市種類[[a,1]],"國中各校",科目種類.idc[[b,1]],"學力表現", sep = ""),
               x = "學校參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          theme_minimal() + thm2()
        
        
        ## 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = ms.p2,
                          location = ph_location_fullsize() )
      }
    }
  }
  
  # 縣市 "各區域" 科目 "學力表現" (折線圖)
  for(c in 1:nrow( 幾年級.idc ) )  {
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[c,1]] ]
    # 指定 科目 for h
    科目種類.idcs    <- df.all.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
    
    
    df.all.idcy[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.all.idcy[["地區"]] )
    # 折線圖資料 ( 要用 )
    g4.01 <- df.all.idcy    %>%
      group_by(地區,科目) %>%
      summarise(人數 = n(), 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                  標準差 = round(sd(總平均)+0.0000000000000001,2) , .groups = 'drop_last' ) %>% setorder(人數)
    
    
    g4.01$地區 <- factor(g4.01$地區, levels=unique(g4.01$地區))
    nl = length(unique(g4.01$地區))
    # 確認 "地區" 數量是否 超過 預設範圍
    g4.text          = 4
    if(nrow(g4.01)<25){ g.axis.text.x = 14 } else {g.axis.text.x = 12}
    g4.plot.title    = 18
    g4.axis.title.xy = 16
    g4.legend.title  = 16
    g4.legend.text   = 12
    g4.strip.text.x  = 12
    x.angle = 0
    thm4 <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                             axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                             strip.text.x = element_text(size = g4.strip.text.x ,family = "A"),
                             plot.title   = element_text(hjust = 0.5,size= g4.plot.title ,family = "A"),
                             axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                             axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                             legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                             legend.text  = element_text(colour="black", size = g4.legend.text,family = "A"),# Windows user
                             legend.position="bottom") 
    if (nrow(g4.01) <= 13 ){
      #p0.sc      <- 20
      g4.width   <- 26
      g4.height  <- 13
    }else{if(nrow(g4.01)<=20){
      #p0.sc      <- 28
      g4.width   <- 40
      g4.height  <- 18
    }else{if(nrow(g4.01) <= 30){
      #p0.sc      <- 36
      g4.width   <- 56
      g4.height  <- 20
    }else{
      #p0.sc      <- 45
      g4.width   <- 70
      g4.height  <- 25
    }            }          }
    
    
    p4<-ggplot(g4.01,aes(x=地區,y=平均數,group=科目,color=科目,label=平均數))+
      geom_line(size=1)+
      geom_point(size=1)+
      # geom_text_repel(size=4,color="black")+
      labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[c,1]],  "各區域整體學力表現", sep = ""),
           x = "區域參與人數由少➜多" ,
           y = "平\n均\n數",family = "A")+
      theme_minimal() + thm4()
    
    
    # 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = p4,
                      location = ph_location_fullsize() )
    
    for(d in 1:nrow( 科目種類.idcs )){
      #  開始前 重置
      df.single <- data.table() 
      df.single.path <-paste( old.path,"/",科目種類.idcs[[d,1]],幾年級.idc[[c,1]],".csv" ,  sep="")
      # 讀入 單 科目*年級 檔案   
      df.single   <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
      # 指定 單檔案  的縣市
      df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
      
      df.single.c[["地區"]] <- gsub("*",  paste( "\n",  sep=""),df.single.c[["地區"]] )
      
      
      #第二部分(各區域):
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
        names(soc.t3.order) <- gsub("縣市$",  paste( 科目種類.idcs[[d,1]],"-縣市人數",  sep=""),names(soc.t3.order))
        names(soc.p3.order) <- gsub("縣市$",  paste( 科目種類.idcs[[d,1]],"-縣市人數",  sep=""),names(soc.p3.order))
        # 將 各向度 通過率 寬轉長 表格 (畫圖用)
        #1. 總答對率(低到高)排序
        soc.t.22<-pivot_longer(soc.t2.order, cols = -地區, names_to = "向度", values_to = "平均數"  )
        #2. 區域人數(少到多)排序
        soc.p.22<-pivot_longer(soc.p2.order, cols = -地區, names_to = "向度", values_to = "平均數"  )
        
        
        # 在地區前面 加上 "該科目" 文字
        #1. 總答對率(低到高)排序
        names(soc.t2.order) <- gsub("地區$",  paste( 科目種類.idcs[[d,1]],"-地區",  sep=""),names(soc.t2.order))
        #2. 區域人數(少到多)排序
        names(soc.p2.order) <- gsub("地區$",  paste( 科目種類.idcs[[d,1]],"-地區",  sep=""),names(soc.p2.order))
        
        
        # 確認 "地區" 數量是否 超過 預設範圍
        if (nrow(soc.s.21) <= 13 ){
          g4.width   <- 26
          g4.height  <- 13
        }else{if(nrow(soc.s.21)<=20){
          g4.width   <- 40
          g4.height  <- 18
        }else{if(nrow(soc.s.21) <= 30){
          g4.width   <- 56
          g4.height  <- 20
        }else{
          g4.width   <- 70
          g4.height  <- 25
        }            }          }
        
      }
      # 縣市 "各區域" 科目 "學力表現" (折線圖) + "各校"
      if(trun.2 == "1"){
        n.max = max(nchar(as.character(unique(names(soc.s.21)))) , na.rm = TRUE)
        nl    = unique(names(soc.s.21)) %>%length()
        g4.text          = 4
        if(nrow(soc.s.21)<25){ g.axis.text.x = 14 } else {g.axis.text.x = 12}
        g4.plot.title    = 18
        g4.axis.title.xy = 16
        g4.legend.title  = 16
        if(n.max<6){ g4.legend.text = 12 } else {g4.legend.text = 9}
        g4.strip.text.x  = 12
        x.angle = 0
        if(nl<=10 & n.max<=5){gnrow = 1 } else{ if (nl<=10 | n.max<=8){gnrow = 2 } else{gnrow = 3}}
        thm4 <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                                 axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                 strip.text.x = element_text(size = g4.strip.text.x ,family = "A"),
                                 plot.title   = element_text(hjust = 0.5,size= g4.plot.title ,family = "A"),
                                 axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                                 axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                                 legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                                 legend.text  = element_text(colour="black", size = g4.legend.text,family = "A"),# Windows user
                                 legend.position="bottom"
        ) 
        #1. 總答對率(低到高)排序
        
        
        p5<-ggplot(soc.t.22,aes(x=地區,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[c,1]],科目種類.idcs[[d,1]],  "各區域學力表現", sep = ""),
               x = paste( 科目種類.idcs[[d,1]],  "區域平均答對率由低➜高", sep = ""),
               y = "平\n均\n數",family = "A")+
          guides( col   = guide_legend  (nrow  = gnrow  )  )+
          theme_minimal() + thm4()
        
        
        
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = p5,
                          location = ph_location_fullsize() )
        
        
        #2. 區域人數(少到多)排序
        p6<-ggplot(soc.p.22,aes(x=地區,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[c,1]],科目種類.idcs[[d,1]],  "各區域學力表現", sep = ""),
               x = "區域參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          guides( col = guide_legend  (nrow  = gnrow ) )+
          theme_minimal() + thm4()
        
        # 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = p6,
                          location = ph_location_fullsize() )
        
        
        #第二部分(各校):
        { thm5 <- function() theme(axis.text.x  = element_text(colour="white",angle = 20, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                                   axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                   strip.text.x = element_text(size = g4.strip.text.x ,family = "A"),
                                   plot.title   = element_text(hjust = 0.5,size= g4.plot.title ,family = "A"),
                                   axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                                   axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                                   legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                                   legend.text  = element_text(colour="black", size = g4.legend.text,family = "A"),# Windows user
                                   legend.position="bottom") 
          # 指定 "學校ID" 跟 各向度
          df.single.c.school <- df.single.c[,c(2,4,21:ncol(df.single.c)),with = FALSE]
          # 各向度平均數
          soc.school.21.m <- df.single.c.school    %>%
            group_by(學校ID)  %>%
            summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last') 
          # 人數
          soc.school.21.p <- df.single.c.school    %>%
            group_by(縣市,學校ID)  %>%
            summarise(人數 = n(), .groups = 'drop_last' )   %>%
            setorder(人數)
          # 合併 人數 * 各向度平均 (原始)
          soc.school.21 <- full_join(soc.school.21.p,soc.school.21.m,by = "學校ID") %>% data.table()
          
          
          #  表格用
          #1. 總答對率(低到高)排序
          soc.school.t.order  <- setorder(soc.school.21,總平均)
          soc.school.t.order$學校ID <- factor(soc.school.t.order$學校ID, levels=unique(soc.school.t.order$學校ID))
          soc.school.t2.order <- soc.school.t.order[,c(-1,-3:-4)]
          
          
          #2. 區域人數(少到多)排序
          soc.school.p.order<-setorder(soc.school.21,人數)
          soc.school.p.order$學校ID <- factor(soc.school.p.order$學校ID, levels=unique(soc.school.p.order$學校ID))
          soc.school.p2.order <- soc.school.p.order[,c(-1,-3:-4)]
          
          
          
          # 將 各向度 通過率 寬轉長 表格 (畫圖用)
          #1. 總答對率(低到高)排序
          soc.school.t.22<-pivot_longer(soc.school.t2.order, cols = -學校ID, names_to = "向度", values_to = "平均數"  )
          #2. 區域人數(少到多)排序
          soc.school.p.22<-pivot_longer(soc.school.p2.order, cols = -學校ID, names_to = "向度", values_to = "平均數"  )
          
          
        }
        
        
        #1. 總答對率(低到高)排序
        p7<-ggplot(soc.school.t.22,aes(x=學校ID,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=0.5)+
          geom_point(size=0.5)+
          xlab(NULL)+
          # geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[c,1]],科目種類.idcs[[d,1]],  "各校學力表現", sep = ""),
               x = paste( 科目種類.idcs[[d,1]],  "學校平均答對率由低➜高", sep = ""),
               y = "平\n均\n數",family = "A")+
          guides( col = guide_legend  (nrow  = gnrow ) )+
          theme_minimal() + thm5()
        
        # 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = p7,
                          location = ph_location_fullsize() )
        
        #2. 區域人數(少到多)排序
        p8<-ggplot(soc.school.p.22,aes(x=學校ID,y=平均數,group=向度,color=向度,label=平均數))+
          geom_line(size=0.5)+
          geom_point(size=0.5)+
          xlab(NULL)+
          #geom_text_repel(size=4,color="black")+
          labs(title = paste( 縣市種類[[a,1]],幾年級.idc[[c,1]],科目種類.idcs[[d,1]],  "各校學力表現", sep = ""),
               x = "學校參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          guides( col = guide_legend  (nrow  = gnrow ) )+
          theme_minimal() + thm5()
        
        
        # 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = p8,
                          location = ph_location_fullsize() )
      }
      
      
      #  縣市 年級 科目 "學力表現" (長條圖)
      if (trun.3 == "1"){    
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
        sub.names<-paste(科目種類.idcs[[d,1]],幾年級.idc[[c,1]],"向度名稱" ,  sep="")
        hl.sc.ctp<-pivot_longer(hl.sc.ct4, cols = - 計算類別, names_to = sub.names, values_to = "平均數"  )
        hl.sc.ct5<-pivot_wider(hl.sc.ctp,names_from = 計算類別, values_from = 平均數)
        #  設定 變數 的排序順序
        hl.sc.ctp[[sub.names]] <- factor(hl.sc.ctp[[sub.names]],levels = unique(hl.sc.ct5[[sub.names]]) )
        
        
        
        #繪圖設定
        n.max = max(nchar(as.character(unique(hl.sc.ctp[[sub.names]]))) , na.rm = TRUE)
        nl    = unique(hl.sc.ctp[[sub.names]]) %>%length()
        if(nl<=6 & n.max<= 6){g3.text = 5 }else{if(nl<=8 & n.max<= 8){g3.text = 3} else{g3.text = 2}}
        if(nl<=6 | n.max<= 6){g3.axis.text.x = 16 }else{if(nl<=8 & n.max<= 8){g3.text = g3.axis.text.x =10} else{g3.axis.text.x = 8}}
        g3.width         = 22
        g3.height        = 11
        g3.plot.title    = 18
        g3.axis.title.xy = 16
        g3.legend.title  = 16
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
                                 legend.text  = element_text(colour="black", size = g3.legend.text,family = "A"),# Windows user
                                 legend.position="bottom") 
        # 長條圖
        # geom_text_repel(hjust= 0.5 , vjust = 1.6,color="white",
        #                 position = position_dodge(0.9), size=g1.text)
        p9 <- ggplot(data=hl.sc.ctp, aes_string(x=`sub.names`, y= "平均數"  ,fill="計算類別",label="平均數" )) +
          geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
          geom_text(hjust= 0.5 , vjust = 1.6,color="white",
                    position = position_dodge(0.9), size=g3.text)+
          scale_fill_manual(values=c("#2b47d2","#71ccfe","#fe0733" ))+
          labs(title = paste( 縣市種類[[a,1]],科目種類.idcs[[d,1]],幾年級.idc[[c,1]],"學力表現", sep = ""),
               x = "向度名稱",
               y = "平\n均\n數",family = "A")+
          guides(    color = guide_colorbar(order = 0  ),
                     fill  = guide_legend  (order = 1  ) )+
          theme_minimal() + thm3()
        
        # 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = p9,
                          location = ph_location_fullsize() )
      }
    }
  }  
  
  # 不同家庭背景 (長條圖)
  for(e in 1:nrow( 幾年級.idc ) )  { # 1:nrow( 幾年級.idc )
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[e,1]] ]
    
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
    
    
    
    
    
    if (trun.4 == "1"){
      #繪圖設定
      g1.text          = 5
      g1.axis.text.x   = 14
      g1.width         = 22
      g1.height        = 11
      g1.plot.title    = 18
      g1.axis.title.xy = 16
      g1.legend.title  = 16
      g1.legend.text   = 12
      g1.strip.text.x  = 12
      thm <- function() theme(axis.text.x  = element_text(angle = 0, vjust = 0.5,size= g1.axis.text.x ,family = "A"),
                              axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                              strip.text.x = element_text(size = g1.strip.text.x ,family = "A"),
                              plot.title   = element_text(hjust = 0.5,size= g1.plot.title ,family = "A"),
                              axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g1.axis.title.xy,family = "A"),
                              axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g1.axis.title.xy,family = "A"),
                              legend.title = element_text(colour="black", size=g1.legend.title, face="bold",family = "A"),
                              legend.text  = element_text(colour="black", size = g1.legend.text,family = "A"),# Windows user
                              legend.position="bottom") 
      # 長條圖
      p10 <- ggplot(data=df.Identity.2, aes(x=身份, y= 總平均數  ,fill=科目,label=總平均數 )) +
        geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
        geom_text(hjust= 0.5 , vjust = 1.6,color="white",
                  position = position_dodge(0.9), size=g1.text)+
        # scale_fill_manual(values=c("#2b5ffe","#feed4d","#82f566","f74b66" ))+
        labs(title = paste( 縣市種類[[a,1]],"不同家庭背景",幾年級.idc[[e,1]],"學生學力表現", sep = ""),
             y = "平\n均\n數",family = "A")+
        guides(    color = guide_colorbar(order = 0),
                   fill = guide_legend(order = 1)  )+
        theme_minimal() + thm()
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = p10,
                        location = ph_location_fullsize() )
      
    }
  } 
  
  # 各卷人數與成績相關表格 (表)
  for(g in 1:nrow( 幾年級.idc ) )  {
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[g,1]] ]
    # 指定 科目 for h
    科目種類.idcs    <- df.all.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
    
    
    for(h in 1:nrow( 科目種類.idcs )){
      #  開始前 重置
      df.single <- data.table()
      df.single.path <- paste( old.path,"/",科目種類.idcs[[h,1]],幾年級.idc[[g,1]],".csv" ,  sep="")
      # 讀入 單 科目*年級 檔案   
      df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
      
      # 指定 單檔案  的縣市
      df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
      
      # 排除 計算不需要的 欄位
      df.single.cs  <- df.single.c[,c(-1:-4,-6:-20)]
      
      
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
                             paste( 幾年級.idc[[g,1]], 科目種類.idcs[[h,1]] ,"-Pearson相關",  sep=""),names(cor.ct3))
      
      # 將 相關表 寬轉長 表格 
      cor.ct3.L <- pivot_longer(cor.ct3,
                                cols = -paste( 幾年級.idc[[g,1]], 科目種類.idcs[[h,1]] ,"-Pearson相關",  sep=""),
                                names_to = "向度", values_to = "平均數"   )
      # 將 各地區 人數 長轉寬 表格 #spread(df.Identity.1, key = 科目  , value = 人數)
      cor.ct3.W <- pivot_wider(cor.ct3.L , names_from = paste( 幾年級.idc[[g,1]], 科目種類.idcs[[h,1]] ,"-Pearson相關",  sep=""),
                               values_from = 平均數)
      names(cor.ct3.W) <- c(paste( 幾年級.idc[[g,1]], 科目種類.idcs[[h,1]] ,"-各向度",  sep=""),"各區人數\n答對率相關","各區人數\n標準差相關")
      
      
      
      # 相關表
      if (trun.5 == "1"){
        # 匯入 表格
        lay  <- lay_new(matrix(1:2,nc = 2),widths=c(1,1),heights=c(1))
        titleLay <- lay_new(1, widths = 1, heights = 1)
        layout <- lay_bind_row(titleLay, lay, heights = c(1,7))
        # lay_show(layout)
        ## create officer layout
        offLayout <- phl_layout(layout,
                                margins = c( 0.2 , 0.2, 0.2, 0.2 ),
                                innerMargins = rep(0.15,4))
        
        if( h == 1 | h == 3 ){
          my_ppt <- add_slide(my_ppt) # 新增分頁
          
          
          style_cor <- fpar(ftext("相關表",
                                  fp_text(bold = TRUE, font.family = "標楷體",
                                          font.size = 40, color = "#333333")))
          
          my_ppt <- ph_with(x= my_ppt, value = style_cor,id = 1,location = ph_location_type(type = "title"))
          
        }
        
        # 匯入 表格 (按照 Layout 設定)
        if( h == 1 | h == 2) {id_cor = (h+1)} 
        if( h == 3 | h == 4) {id_cor = (h-1)}
        flTable <- phl_adjust_table(cor.ct3.W, olay = offLayout, id = id_cor )
        flTable <-  bg(flTable, bg = "#add8e6", part = "header")
        flTable <-  bg(flTable, bg = "#ffffff", part = "body")
        flTable <- color(flTable, color = "#000000")
        flTable <- fontsize(flTable, part = "all", size = 12)
        flTable <- width(flTable, width = c(2,1,1))
        flTable <- font(flTable, part = "all" ,fontname = "標楷體")
        
        my_ppt <- phl_with_flextable(my_ppt, olay = offLayout, id_cor, flTable)
        
      }
      
    }}
  
  # 各卷 學力表現 (表)
  for(k in 1:nrow( 幾年級.idc ) )  {
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[k,1]] ]
    # 指定 科目 for h
    科目種類.idcs    <- df.all.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
    
    for(m in 1:nrow( 科目種類.idcs )){
      #  開始前 重置
      df.single <- data.table() 
      df.single.path <-paste( old.path,"/",科目種類.idcs[[m,1]],幾年級.idc[[k,1]],".csv" ,  sep="")
      # 讀入 單 科目*年級 檔案   
      df.single   <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
      # 指定 單檔案  的縣市
      df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]  
      #  縣市 年級 科目 "學力表現" (長條圖)
      if (trun.3 == "1"){    
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
        # 合併表格 並 寬轉長 # ,hl.sc.ct3
        hl.sc.ct4 <- bind_rows(hl.sc.ct1,hl.sc.ct2) 
        hl.sc.ct4$計算類別 <- factor(hl.sc.ct4$計算類別, levels=c("平均","標準差","高分組-低分組")) 
        # 將 各向度 通過率 寬轉長 表格
        sub.names<-paste(科目種類.idcs[[m,1]],幾年級.idc[[k,1]],"向度名稱" ,  sep="")
        hl.sc.ctp<-pivot_longer(hl.sc.ct4, cols = - 計算類別, names_to = sub.names, values_to = "平均數"  )
        hl.sc.ct5<-pivot_wider(hl.sc.ctp,names_from = 計算類別, values_from = 平均數)
        #  設定 變數 的排序順序
        hl.sc.ct5[[sub.names]] <- factor(hl.sc.ct5[[sub.names]],levels = unique(hl.sc.ct5[[sub.names]]) )
        
        
        
        # 匯入 表格
        lay  <- lay_new(matrix(1:4,nc = 2),widths=c(1,1),heights=c(1,1))
        titleLay <- lay_new(1, widths = 1, heights = 1)
        layout <- lay_bind_row(titleLay, lay, heights = c(1,7))
        # lay_show(layout)
        offLayout <- phl_layout(layout,
                                margins = c( 0.2 , 0.2, 0.2, 0.2 ),
                                innerMargins = rep(0.15 , 4))
        
        if( m == 1 ){ 
          
          my_ppt <- add_slide(my_ppt)  # 新增分頁
          style_cor <- fpar( ftext( paste(縣市種類[[a,1]],幾年級.idc[[k,1]],"各科學力表現", sep=""),
                                    fp_text(bold = TRUE, font.family = "標楷體",
                                            font.size = 40, color = "#333333")))
          
          my_ppt <- ph_with(x= my_ppt, value = style_cor,id = 1,location = ph_location_type(type = "title"))
        } 
        
        flTable <- phl_adjust_table(hl.sc.ct5, olay = offLayout, id = (m+1) )
        flTable <-  bg(flTable, bg = "#add8e6", part = "header")
        flTable <-  bg(flTable, bg = "#ffffff", part = "body")
        flTable <- color(flTable, color = "#000000")
        flTable <- fontsize(flTable, part = "all", size = 12)
        flTable <- width(flTable, width = c(3,0.7,0.7))
        flTable <- font(flTable, part = "all" ,fontname = "標楷體")
        
        my_ppt <- phl_with_flextable(my_ppt, olay = offLayout, (m+1), flTable)
        
        
        
        
      }
      
    }
  }
  
  # 各區 平均學力 (表)
  for(g in 1:nrow( 幾年級.idc ) )  {
    # 指定 年級
    df.all.idcy <- df.all.idc[年級 == 幾年級.idc[[g,1]] ]
    # 指定 科目 for h
    科目種類.idcs    <- df.all.idcy[,.(N = .N),by = .(科目)] %>% setorder(科目)
    
    
    for(h in 1:nrow( 科目種類.idcs )){
      #  開始前 重置
      df.single <- data.table()
      df.single.path <- paste( old.path,"/",科目種類.idcs[[h,1]],幾年級.idc[[g,1]],".csv" ,  sep="")
      # 讀入 單 科目*年級 檔案   
      df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
      
      # 指定 單檔案  的縣市
      df.single.c <- df.single[縣市 == 縣市種類[[a,1]] ]
      # 計算 各地區 人數、平均、標準差
      soc.s.11 <- df.single.c    %>%
        group_by(地區)  %>%
        summarise(人數 = n(), 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                    標準差 = round(sd(總平均)+0.0000000000000001,2) , .groups = 'drop_last' )
      # 在地區前面 加上 "該科目" 文字
      names(soc.s.11) <- gsub("地區$",
                              paste( 科目種類.idcs[[h,1]],"-地區",  sep=""),names(soc.s.11))
      
      
      
      
      
      
      
      
      # 各區平均學力 表
      if (trun.5 == "1"){
        # 匯入 表格
        lay  <- lay_new(matrix(1:2,nc = 2),widths=c(1,1),heights=c(1))
        titleLay <- lay_new(1, widths = 1, heights = 1)
        layout <- lay_bind_row(titleLay, lay, heights = c(1,7))
        # lay_show(layout)
        ## create officer layout
        offLayout <- phl_layout(layout,
                                margins = c( 0.2 , 0.2, 0.2, 0.2 ),
                                innerMargins = rep(0.15,4))
        
        if( h == 1 | h == 3 ){
          my_ppt <- add_slide(my_ppt) # 新增分頁
          
          
          style_cor <- fpar( ftext( paste(縣市種類[[a,1]],幾年級.idc[[g,1]],"各區平均學力", sep="") ,
                                    fp_text(bold = TRUE, font.family = "標楷體",
                                            font.size = 40, color = "#333333")))
          
          my_ppt <- ph_with(x= my_ppt, value = style_cor,id = 1,location = ph_location_type(type = "title"))
          
        }
        
        # 匯入 表格 (按照 Layout 設定)
        if( h == 1 | h == 2) {id_cor = (h+1)} 
        if( h == 3 | h == 4) {id_cor = (h-1)}
        if( nrow(soc.s.11)<40 ) { f.size = 9 } 
        if( nrow(soc.s.11)<35 ) { f.size = 11 } 
        if( nrow(soc.s.11)<29 ) { f.size = 12 } 
        if( nrow(soc.s.11)<21 ) { f.size = 13 } 
        if( nrow(soc.s.11)<14 ) { f.size = 14 } 
        
        flTable <- phl_adjust_table(soc.s.11, olay = offLayout, id = id_cor )
        flTable <-  bg(flTable, bg = "#add8e6", part = "header")
        flTable <-  bg(flTable, bg = "#ffffff", part = "body")
        flTable <- color(flTable, color = "#000000")
        flTable <- fontsize(flTable, part = "all", size = f.size)
        flTable <- width(flTable, width = c(1.5,1,1,1))
        flTable <- font(flTable, part = "all" ,fontname = "標楷體")
        
        my_ppt <- phl_with_flextable(my_ppt, olay = offLayout, id_cor, flTable)
        
      }
      
    }}  
  
  
  if( nrow( df.CFA.p [縣市 == 縣市種類[[a,1]] ]  ) > 5  ){
    # CFA 問卷 與 成績 相關 (長條圖)  
    for(i in 1:nrow(年級科目種類.CFA.idc) )  {
      df.CFA.p.idcys <- df.CFA.p.idc[年級 == 年級科目種類.CFA.idc[[i,1]] &  科目 == 年級科目種類.CFA.idc[[i,2]] ] %>% setorder(-`Pearson 相關`)
      
      df.CFA.p.idcys$`向度`<- factor(df.CFA.p.idcys$`向度`, levels=unique(df.CFA.p.idcys$`向度`))
      
      #繪圖設定
      g11.text          = 8
      g11.axis.text.x   = 14
      g11.width         = 22
      g11.height        = 11
      g11.plot.title    = 18
      g11.axis.title.xy = 16
      g11.legend.title  = 16
      g11.legend.text   = 12
      g11.strip.text.x  = 12
      thm11 <- function() theme(axis.text.x  = element_text(angle = 20, vjust = 0.5,size= g1.axis.text.x ,family = "A"),
                                axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                                strip.text.x = element_text(size = g1.strip.text.x ,family = "A"),
                                plot.title   = element_text(hjust = 0.5,size= g1.plot.title ,family = "A"),
                                axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g1.axis.title.xy,family = "A"),
                                axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g1.axis.title.xy,family = "A"),
                                legend.title = element_text(colour="black", size=g1.legend.title, face="bold",family = "A"),
                                legend.text  = element_text(colour="black", size = g1.legend.text,family = "A"),# Windows user
                                legend.position="bottom") 
      # 長條圖
      p11 <- ggplot(data=df.CFA.p.idcys, aes(x=`向度`, y= `Pearson 相關`  ,label=`Pearson 相關` )) +
        geom_bar(fill="#b8e6fa",stat="identity", position=position_dodge(),size=.3)+
        geom_text(hjust= 0.5 , vjust = 1.6,color="#000000",
                  position = position_dodge(0.9), size=g1.text)+
        labs(title = paste( 縣市種類[[a,1]],年級科目種類.CFA.idc[[i,1]],"學生個人變項與",年級科目種類.CFA.idc[[i,2]],"學力表現之相關", sep = ""),
             x = NULL,
             y = "平\n均\n數",family = "A")+
        guides(    color = guide_colorbar(order = 0),
                   fill = guide_legend(order = 1)  )+
        theme_minimal() + thm11()
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = p11,
                        location = ph_location_fullsize() )
      
    }
  }
  
  
  
  
  
  
  
  
  print(my_ppt, target = paste("C:/Users/user/Desktop/測試/",縣市種類[[a,1]],"-1020.pptx", sep=""))
}














