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
  
  ########## 將 身分 合併進 資料
  df.all.id<- data.table(df.all[,1:20],身份,df.all[,21])
  
  # 讀入 CFA 總檔案
  old.CFA.path<- "C:/Users/user/Documents/台中教育大學/學力工作/20200918-學力報告製作/問卷資料"
  new.CFA.path<-list.files(path = old.CFA.path,full.names = T)#指定資料夾內的資料夾清單
  df.CFA <- read_excel(new.CFA.path[1] , sheet = 1 , col_names = TRUE , na = "" ) %>% data.table()
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
  # df.CFA.p$縣市 <- factor(df.CFA.p$縣市, levels=unique(df.all$縣市))
  
  
  ##########  輸出表格設定 1
  hs1 <- createStyle(fontColour = "#000000", fgFill = "#ffffff",
                     halign = "center", valign = "center", textDecoration = "bold",
                     border = "TopBottomLeftRight")
  # 表格字體設定
  windowsFonts(A=windowsFont("標楷體"))
  # gridExtra 表格(圖) 字體設定 # https://www.rdocumentation.org/packages/gridExtra/versions/2.3/topics/tableGrob
  mytheme <- gridExtra::ttheme_default( base_family = "A" )
  
  trun.1 <- "1" # 各縣市 "各區域" 科目 "學力表現" (折線圖)
  trun.2 <- "1" # 各縣市 幾年級  科目  "各區域學力表現" (折線圖)
  trun.3 <- "1" # 各縣市 年級 科目 "學力表現" (長條圖)
  trun.4 <- "1" # 各縣市 不同家庭背景   年級 學生學力表現
  trun.5 <- "1" # 各縣市 年級 科目 相關表
}

#繪圖設定
{
    g.width   <- 28
    g.height  <- 14
    g.text          = 4
    g.axis.text.x   = 14
    g.plot.title    = 18
    g.axis.title.xy = 16
    g.legend.title  = 16
    g.legend.text   = 12
    g.strip.text.x  = 12
    x.angle = 0
    thm <- function() theme(axis.text.x  = element_text(angle = x.angle, vjust = 0.5,size= g.axis.text.x ,family = "A"),
                             axis.text.y  = element_text(angle = 0, vjust = 0.5,size=16,family = "A"),
                             strip.text.x = element_text(size = g.strip.text.x ,family = "A"),
                             plot.title   = element_text(hjust = 0.5,size= g.plot.title ,family = "A"),
                             axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g.axis.title.xy,family = "A"),
                             axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g.axis.title.xy,family = "A"),
                             legend.title = element_text(colour="black", size=g.legend.title, face="bold",family = "A"),
                             legend.text  = element_text(colour="black", size = g.legend.text,family = "A"),# Windows user
                             legend.position="bottom")

}

# 資料整理
{
  # 各縣市 人數
  df.all.people <- df.all.id    %>%
    group_by(縣市)  %>%
    summarise( 人數 = n(), .groups = 'drop_last' )   %>% setorder(人數)    %>% data.table()
  # 繪圖前 調整 縣市 格式(\n)
  df.all.people[["縣市"]] <- gsub("*",  paste( "\n",  sep=""),df.all.people[["縣市"]])
  # 各縣市 各科目 各年級
  df.all.cs <- df.all.id    %>%
    group_by(科目,年級)  %>%
    summarise( 人數 = n(), .groups = 'drop_last' )   %>% setorder(科目,年級)    %>% data.table()


  
  年級科目種類.CFA.id    <- df.CFA.p[,.(N = .N),by = .(年級,科目)] %>% setorder(年級,科目)
  #建立 ppt
  my_ppt <- read_pptx()
}


# 執行流程


# 1. 各縣市 各年級 學力表現 (科目)
for(a in 1 : length(年級排序)){
  # 指定 科目
  df.all.idy <- df.all.id[年級 == 年級排序[a] ]
  
  
  # 平均數
  df.subject <- df.all.idy    %>%
    group_by(縣市,科目) %>%
    summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
              人數 = n(), .groups = 'drop_last' )   %>% data.table()
  # 繪圖前 調整 縣市 格式(\n)
  df.subject[["縣市"]] <- gsub("*",  paste( "\n",  sep=""),df.subject[["縣市"]])
  # 按照 各縣市 人數 排序
  df.subject[["縣市"]] <- factor(df.subject[["縣市"]], levels=unique(df.all.people[["縣市"]]))
  
  df.subject[["科目"]] <- factor(df.subject[["科目"]], levels=unique(科目排序))
  
  # 繪圖
  subject.polt <- ggplot(data=df.subject, aes(x=縣市, y= 平均數  ,group=科目,color=科目,label=平均數 )) +
    geom_line(size=1)+
    geom_point(size=1)+
    # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
    labs(title = paste( "各縣市",年級排序[a],"學力表現", sep = ""),
         x = "縣市參與人數由少➜多" ,
         y = "平\n均\n數",family = "A")+
    theme_minimal() + thm()
  
  ## 匯入圖片
  my_ppt <- add_slide(my_ppt) # 新增分頁
  my_ppt <- ph_with(x = my_ppt, value = subject.polt,
                    location = ph_location_fullsize() )
  
}


# 2. 各縣市 各科目 學力表現 (年級)
for(b in 1 : length(科目排序)){
  # 指定 科目
  df.all.ids <- df.all.id[科目 == 科目排序[b] ]
  # 檢定 縣市數量
  counties.N<-unique(df.all.ids[["縣市"]]) %>% length()
  
  # 平均數
  df.year <- df.all.ids    %>%
    group_by(年級,縣市) %>%
    summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
              人數 = n(), .groups = 'drop_last' )   %>% data.table()
  # 繪圖前 調整 縣市 格式(\n)
  df.year[["縣市"]] <- gsub("*",  paste( "\n",  sep=""),df.year[["縣市"]])
    # 按照 各縣市 人數 排序
  df.year[["縣市"]] <- factor(df.year[["縣市"]], levels=unique(df.all.people[["縣市"]]))

  df.year[["年級"]] <- factor(df.year[["年級"]], levels=unique(年級排序))

  # 國小 1~6 年級
  df.year.es <- filter(df.year,年級 %in% 國小)
  # 國中 7~8 年級
  df.year.ms <- filter(df.year,年級 %in% 國中)
  if(counties.N != "1"){
      
      # 國小 es
      es.polt <- ggplot(data=df.year.es, aes(x=縣市, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
        geom_line(size=1)+
        geom_point(size=1)+
        # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
        labs(title = paste( "各縣市國小",科目排序[b],"學力表現", sep = ""),
             x = "縣市參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        theme_minimal() + thm()
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = es.polt,
                        location = ph_location_fullsize() )
      # 國中 ms
      if(nrow(df.year.ms)>1){
        ms.polt <- ggplot(data=df.year.ms, aes(x=縣市, y= 平均數  ,group=年級,color=年級,label=平均數 )) +
          geom_line(size=1)+
          geom_point(size=1)+
          # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
          labs(title = paste( "各縣市國中",科目排序[b],"學力表現", sep = ""),
               x = "縣市參與人數由少➜多" ,
               y = "平\n均\n數",family = "A")+
          theme_minimal() + thm()
        
        
        ## 匯入圖片
        my_ppt <- add_slide(my_ppt) # 新增分頁
        my_ppt <- ph_with(x = my_ppt, value = ms.polt,
                          location = ph_location_fullsize() )
      }
    } 
  else{
    g.text = 10
    # 國小 es
    es.polt <- ggplot(data=df.year.es, aes(x=縣市, y= 平均數  ,fill=年級,label=平均數 )) +
      geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3,width = 0.5)+
      geom_text(hjust= c(-0.9,1.5) , vjust = 1.6,color="white",
                position = position_dodge(0.9), size=g.text)+
      # scale_fill_manual(values=c("#2b5ffe","#feed4d","#82f566","f74b66" ))+
      labs(title = paste( "各縣市國小",科目排序[b],"學力表現", sep = ""),
           x = "縣市參與人數由少➜多" ,
           y = "平\n均\n數",family = "A")+
      guides(    color = guide_colorbar(order = 0),
                 fill = guide_legend(order = 1)  )+
      theme_minimal() + thm()
  
    ## 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = es.polt,
                      location = ph_location_fullsize() )
    
    # 國中 ms
    if(nrow(df.year.ms)>1){
      ms.polt <- ggplot(data=df.year.ms, aes(x=縣市, y= 平均數  ,fill=年級,label=平均數 )) +
        geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3,width = 0.5)+
        geom_text(hjust= c(-0.9,1.5) , vjust = 1.6,color="white",
                  position = position_dodge(0.9), size=g.text)+
        # scale_fill_manual(values=c("#2b5ffe","#feed4d","#82f566","f74b66" ))+
        labs(title = paste( "各縣市國小",科目排序[b],"學力表現", sep = ""),
             x = "縣市參與人數由少➜多" ,
             y = "平\n均\n數",family = "A")+
        guides(    color = guide_colorbar(order = 0),
                   fill = guide_legend(order = 1)  )+
        theme_minimal() + thm()
      
      
      ## 匯入圖片
      my_ppt <- add_slide(my_ppt) # 新增分頁
      my_ppt <- ph_with(x = my_ppt, value = ms.polt,
                        location = ph_location_fullsize() )
    }
    g.text = 4
  }

  
  
}


# 3. 各縣市 各科目 各年級 學力表現 (向度)
for(c in 1 : nrow(df.all.cs)){

  #  開始前 重置
  df.single <- data.table()
  df.single.path <- paste( old.path,"/",df.all.cs[[c,1]],df.all.cs[[c,2]],".csv" ,  sep="")
  # 讀入 單 科目*年級 檔案   
  df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
  
  
  # 繪圖前 調整 縣市 格式(\n)
  df.single[["縣市"]] <- gsub("*",  paste( "\n",  sep=""),df.single[["縣市"]])

  # 折線圖資料 ( 要用 )
  df.single.order <- df.single    %>%
    group_by(縣市) %>%
    summarise(人數 = n(),
              平均數 = round(mean(總平均)+0.0000000000000001,2)  , .groups = 'drop_last' )
  

  
  # 排除 計算不需要的 欄位
  df.single.cs  <- df.single[,c(-3:-21)]

  
  # 平均數
  df.subscale <- df.single.cs    %>%
    group_by(縣市) %>%
    summarise_if(is.numeric,~round(mean(.)+0.0000000000000001,2), na.rm = FALSE, .groups = 'drop_last')%>%
    data.table()  
  
  
  # 按照 各縣市 人數 排序
  df.single.order.p <- setorder(df.single.order,人數)
  df.subscale[["縣市"]] <- factor(df.subscale[["縣市"]], levels=unique(df.single.order.p[["縣市"]]))

  

  
  
  
  
  # 繪圖設定
  n.max = max(nchar(as.character(unique(names(df.single.cs)))) , na.rm = TRUE)
  nl    = unique(names(df.single.cs)) %>%length()
  if(nl<=8 & n.max<=5){gnrow = 1 } else{ if (nl<=10 | n.max<=8){gnrow = 2 } else{gnrow = 3}}
  if(df.all.cs[[c,1]] == "自然" | df.all.cs[[c,1]] == "英語文"){g.legend.text   = 9} else{g.legend.text   = 11}
  
  
  if(nrow(df.single.order) != "1"){
    
    # 將 向度表 寬轉長 表格 
    df.subscale.PL <- pivot_longer(df.subscale,
                                   cols = -"縣市",
                                   names_to = "向度", values_to = "平均數"   )
    
    # 參與人數
    subscale.p.polt <- ggplot(data=df.subscale.PL, aes(x=縣市, y= 平均數  ,group=向度,color=向度,label=平均數 )) +
      geom_line(size=1)+
      geom_point(size=1)+
      # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
      labs(title = paste( "各縣市",df.all.cs[[c,1]],df.all.cs[[c,2]],"學力表現", sep = ""),
           x = "縣市參與人數由少➜多" ,
           y = "平\n均\n數",family = "A")+
      guides( col   = guide_legend  (nrow  = gnrow  )  )+
      theme_minimal() + thm()
    ## 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = subscale.p.polt,
                      location = ph_location_fullsize() )
    
    # 總平均
    
    # 按照 各縣市 平均數 排序
    df.single.order.s <- setorder(df.single.order,平均數)
    df.subscale[["縣市"]] <- factor(df.subscale[["縣市"]], levels=unique(df.single.order.s[["縣市"]]))
    
    # 將 向度表 寬轉長 表格 
    df.subscale.SL <- pivot_longer(df.subscale,
                                   cols = -"縣市",
                                   names_to = "向度", values_to = "平均數"   )
    
    subscale.s.polt <- ggplot(data=df.subscale.SL, aes(x=縣市, y= 平均數  ,group=向度,color=向度,label=平均數 )) +
      geom_line(size=1)+
      geom_point(size=1)+
      # geom_text_repel(size=4,color="black")+ # ,hjust=0.1, vjust=-1
      labs(title = paste( "各縣市",df.all.cs[[c,1]],df.all.cs[[c,2]],"學力表現", sep = ""),
           x = "縣市平均答對率由低➜高" ,
           y = "平\n均\n數",family = "A")+
      guides( col   = guide_legend  (nrow  = gnrow  )  )+
      theme_minimal() + thm()
    ## 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = subscale.s.polt,
                      location = ph_location_fullsize() )
  }
  else{
    # 將 向度表 寬轉長 表格 
    df.subscale.PL <- pivot_longer(df.subscale,
                                   cols = -"縣市",
                                   names_to = "向度", values_to = "平均數"   )
    
    
    # 參與人數
    subscale.p.polt <- ggplot(data=df.subscale.PL, aes(x=縣市, y= 平均數  ,fill=向度,label=平均數 )) +
      geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
      geom_text(hjust= 0.5 , vjust = 1.6,color="white",
                position = position_dodge(0.9), size=g.text)+
      # scale_fill_manual(values=c("#2b5ffe","#feed4d","#82f566","f74b66" ))+
      labs(title = paste( "各縣市",df.all.cs[[c,1]],df.all.cs[[c,2]],"學力表現", sep = ""),
           x = "縣市參與人數由少➜多" ,
           y = "平\n均\n數",family = "A")+
      guides(    color = guide_colorbar(order = 0),
                 fill = guide_legend(order = 1)  )+
      theme_minimal() + thm()
    
    
    ## 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = subscale.p.polt,
                      location = ph_location_fullsize() )
    

  }

  
}


# 4. 各縣市 各科目 各年級 高低分組 (向度)
for(c in 1 : nrow(df.all.cs)){
  #  開始前 重置
  df.single <- data.table()
  df.single.path <- paste( old.path,"/",df.all.cs[[c,1]],df.all.cs[[c,2]],".csv" ,  sep="")
  # 讀入 單 科目*年級 檔案   
  df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
  
  
  ########## 根據總分排名 新增 高、低分組 欄位
  # 先取得 高、低分組線(27%)
  low  <- quantile(df.single$總平均, probs=0.27)
  high <- quantile(df.single$總平均, probs=0.73)
  高低分組 <- 0
  高低分組[df.single$總平均 <= low] <- "低分組"
  高低分組[df.single$總平均 >  low  &  df.single$總平均 <high] <- "中間組" 
  高低分組[df.single$總平均 >= high] <- "高分組"
  高低分組 <- factor(高低分組, levels=c("高分組","中間組","低分組")) 
  高低分組 <-data.table(高低分組) 
  table(高低分組) 
  ########## 將 高、低分組 合併進 資料
  df.single.hl<- data.table(高低分組,df.single[,21:ncol(df.single)])
  
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
  sub.names<-paste(df.all.cs[[c,1]],df.all.cs[[c,2]],"向度名稱" ,  sep="")
  hl.sc.ctp<-pivot_longer(hl.sc.ct4, cols = - 計算類別, names_to = sub.names, values_to = "平均數"  )
  hl.sc.ct5<-pivot_wider(hl.sc.ctp,names_from = 計算類別, values_from = 平均數)
  #  設定 變數 的排序順序
  hl.sc.ctp[[sub.names]] <- factor(hl.sc.ctp[[sub.names]],levels = unique(hl.sc.ct5[[sub.names]]) )
  
  #繪圖設定
  n.max = max(nchar(as.character(unique(hl.sc.ctp[[sub.names]]))) , na.rm = TRUE)
  nl    = unique(hl.sc.ctp[[sub.names]]) %>%length()
  if(nl<=6 & n.max<= 6){g.text = 5 }else{if(nl<=8 & n.max<= 8){g.text = 3} else{g.text = 2}}
  if(nl<=6 | n.max<= 6){g.axis.text.x = 16 }else{if(nl<=8 & n.max<= 8){g.text = g.axis.text.x =10} else{g.axis.text.x = 8}}
  x.angle = 20
  
  
  hl.sc.plot<-ggplot(data=hl.sc.ctp, aes_string(x=`sub.names`, y= "平均數"  ,fill="計算類別",label="平均數" )) +
    geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
    geom_text(hjust= 0.5 , vjust = 1.6,color="white",
              position = position_dodge(0.9), size=g.text)+
    scale_fill_manual(values=c("#2b47d2","#71ccfe","#fe0733" ))+
    labs(title = paste( df.all.cs[[c,1]],df.all.cs[[c,2]],"整體學力表現", sep = ""),
         x = "向度名稱",
         y = "平\n均\n數",family = "A")+
    guides(    color = guide_colorbar(order = 0  ),
               fill  = guide_legend  (order = 1  ) )+
    theme_minimal() + thm()
  
  # 匯入圖片
  my_ppt <- add_slide(my_ppt) # 新增分頁
  my_ppt <- ph_with(x = my_ppt, value = hl.sc.plot,
                    location = ph_location_fullsize() )
  
  x.angle = 0
  g.text  = 4
  g.axis.text.x   = 14
  

}


# 5. 各縣市 各年級 各科目 不同身分 (科目)
for(a in 1 : length(年級排序)){
  # 指定 科目
  df.all.idy <- df.all.id[年級 == 年級排序[a] ]
  
  #1 人數
  df.Identity.1 <- df.all.idy   %>%
    group_by(科目,身份) %>%
    summarise(人數 = n(), .groups = 'drop_last' )   %>% data.table()
  # 在科目尾部 加上 "-人數" 文字
  df.Identity.1$科目 <- gsub("$", "-人數",df.Identity.1[,科目])
  # 將 各地區 人數 長轉寬 表格 #spread(df.Identity.1, key = 科目  , value = 人數)
  df.Identity.1w <- pivot_wider(df.Identity.1,names_from = 科目, values_from = 人數)
  
  
  
  #2 平均數
  df.Identity.2 <- df.all.idy    %>%
    group_by(科目,身份) %>%
    summarise( 平均數 = round(mean(總平均)+0.0000000000000001,2) , .groups = 'drop_last' )   %>% data.table()
  
  
  
  g.text          = 5
  # 長條圖
  Identity.plot <- ggplot(data=df.Identity.2, aes(x=身份, y= 平均數  ,fill=科目,label=平均數 )) +
    geom_bar(colour="black",stat="identity", position=position_dodge(),size=.3)+
    geom_text(hjust= 0.5 , vjust = 1.6,color="white",
              position = position_dodge(0.9), size=g.text)+
    # scale_fill_manual(values=c("#2b5ffe","#feed4d","#82f566","f74b66" ))+
    labs(title = paste( "不同家庭背景",年級排序[a],"學生整體學力表現", sep = ""),
         y = "平\n均\n數",family = "A")+
    guides(    color = guide_colorbar(order = 0),
               fill = guide_legend(order = 1)  )+
    theme_minimal() + thm()
  
  ## 匯入圖片
  my_ppt <- add_slide(my_ppt) # 新增分頁
  my_ppt <- ph_with(x = my_ppt, value = Identity.plot,
                    location = ph_location_fullsize() )
  
  
  
  
}


# 6. 各縣市 各年級 各科目 學力表現 (科目表)
for(a in 1 : length(年級排序)){
  # 指定 年級
  df.all.idy <- df.all.id[年級 == 年級排序[a] ]
  # 指定 科目 for h
  科目種類.idcs    <- df.all.idy[,.(N = .N),by = .(科目)] %>% setorder(科目)
  
  
  for(h in 1:nrow( 科目種類.idcs )){
  #  開始前 重置
  df.single <- data.table()
  df.single.path <- paste( old.path,"/",科目種類.idcs[[h,1]],年級排序[a],".csv" ,  sep="")
  # 讀入 單 科目*年級 檔案   
  df.single <- fread(df.single.path,na.strings = "",encoding = "UTF-8")
  


  # 計算 各地區 人數、平均、標準差
  soc.s.table <- df.single    %>%
    group_by(縣市)  %>%
    summarise(人數 = n(), 平均數 = round(mean(總平均)+0.0000000000000001,2) ,
                標準差 = round(sd(總平均)+0.0000000000000001,2) , .groups = 'drop_last' )
  # 在地區前面 加上 "該科目" 文字
  names(soc.s.table) <- gsub("縣市$",
                          paste( 科目種類.idcs[[h,1]],"-縣市",  sep=""),names(soc.s.table))
  
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
    
    
    style_cor <- fpar( ftext( paste(年級排序[a],"各縣市平均學力", sep="") ,
                              fp_text(bold = TRUE, font.family = "標楷體",
                                      font.size = 40, color = "#333333")))
    
    my_ppt <- ph_with(x= my_ppt, value = style_cor,id = 1,location = ph_location_type(type = "title"))
    
  }
  
  # 匯入 表格 (按照 Layout 設定)
  if( h == 1 | h == 2) {id_cor = (h+1)} 
  if( h == 3 | h == 4) {id_cor = (h-1)}
  if( nrow(soc.s.table)<40 ) { f.size = 9 } 
  if( nrow(soc.s.table)<35 ) { f.size = 11 } 
  if( nrow(soc.s.table)<29 ) { f.size = 12 } 
  if( nrow(soc.s.table)<21 ) { f.size = 13 } 
  if( nrow(soc.s.table)<14 ) { f.size = 14 } 
  
  flTable <- phl_adjust_table(soc.s.table, olay = offLayout, id = id_cor )
  flTable <-  bg(flTable, bg = "#add8e6", part = "header")
  flTable <-  bg(flTable, bg = "#ffffff", part = "body")
  flTable <- color(flTable, color = "#000000")
  flTable <- fontsize(flTable, part = "all", size = f.size)
  flTable <- width(flTable, width = c(1.5,1,1,1))
  flTable <- font(flTable, part = "all" ,fontname = "標楷體")
  
  my_ppt <- phl_with_flextable(my_ppt, olay = offLayout, id_cor, flTable)
  

  }}


# 7. CFA 問卷 與 成績 相關 (長條圖) # 1 : nrow(年級科目種類.CFA.id)
for(i in 1 : nrow(年級科目種類.CFA.id) )  {
  
    df.CFA.p.ys <- df.CFA.p[年級 == 年級科目種類.CFA.id[[i,1]] &  科目 == 年級科目種類.CFA.id[[i,2]] ] %>%
                    setorder(-`Pearson 相關`)
    
    df.CFA.p.ys[["向度"]]<- factor(df.CFA.p.ys[["向度"]], levels=unique(df.CFA.p.ys[["向度"]]))
    

    x.angle = 20
    # 長條圖
    CFA.plot <-  ggplot(data=df.CFA.p.ys, aes(x=`向度`, y= `Pearson 相關`  ,label=`Pearson 相關` )) +
                 geom_bar(fill="#b8e6fa",stat="identity", position=position_dodge(),size=.3)+
                 geom_text(hjust= 0.5 , vjust = 1.6,color="#000000",
                           position = position_dodge(0.9), size=g.text)+
                 labs(title = paste( 年級科目種類.CFA.id[[i,1]],"整體學生個人變項與",年級科目種類.CFA.id[[i,2]],"學力表現之相關", sep = ""),
                      x = NULL,
                      y = "平\n均\n數",family = "A")+
                 guides(    color = guide_colorbar(order = 0),
                            fill = guide_legend(order = 1)  )+
                 theme_minimal() + thm()
    
    ## 匯入圖片
    my_ppt <- add_slide(my_ppt) # 新增分頁
    my_ppt <- ph_with(x = my_ppt, value = CFA.plot,
                      location = ph_location_fullsize() )
    x.angle = 0
    g.text  = 4
    g.axis.text.x   = 14
  }




print(my_ppt, target = paste("C:/Users/user/Desktop/測試/","整體學力","-1021.pptx", sep=""))

