library(rio)
library(dplyr)
library(openxlsx)
library(dummies)
library(readxl)
library(data.table)
library(ggplot2)
library(tidyverse)
library(ggthemes)

school.score.plot.cc.2<-function(data.path,data,soc.all,plot.turn ) {
  # 第一步 : 讀入
  # 建立 workbook
  wb <- createWorkbook()
  ## 取得 檔名
  dn<-data.path %>% basename() %>% strsplit( "[,. ]") 
  school.name<-dn[[1]][1]
  # 固定 班級為類別
  data$班級<- as.factor(data$班級) 

  ## 取得 分頁名單
  names.files<-getSheetNames(data.path)
  # 取得 分頁數量
  files <- length(names.files)
  for (i in 1:files){
    ## 新增 對應分頁
    addWorksheet(wb, sheetName = names.files[i])
    ## 匯入 對應data
    df <- readWorkbook(data.path,sheet = i)
    writeData(wb, sheet = names.files[i], x = df)
  }
  #############################################################################
  # 畫圖
  windowsFonts(A=windowsFont("標楷體"))

  # 排除 座號
  soc.s<-data[,-3]
  # 計算 各班 各向度 通過率
  soc.1<-soc.s   %>%
    group_by(班級) %>%
    summarise_if(is.numeric,mean, na.rm = TRUE)
  
  # 匯入 班級 + 校 + 全體考生 各向度 通過率
  # soc.all<-read_excel(df.path,sheet = "各班成績（校）") 
  soc.all<- soc.all
  # 單獨 校成績 + 全體考生
  # 計算 班級數量
  cn<-nrow(soc.1)
  # 計算 向度數量(含總平均) 扣除 班級欄位 1欄
  subscale <- ncol(soc.all)-1
  # 計算 向度最長字數
  nl<-soc.all %>% names() %>% nchar() %>% max()
  
  #排除 班級的成績
  soc.sall1<-soc.all[-1:-cn,]
  soc.sall2<-soc.sall1 %>% 
    rename(
      群體 = 班級
    )
  # 將 各向度 通過率 寬轉長 表格
  soc.sall3<-gather(soc.sall2, key = 向度, value = 平均,-群體)
  # 直接匯入
  socs<-gather(soc.all, key = 向度, value = 平均,-班級)
  socs[,3]<-round(x = socs[,3],digits = 2)
  # 將 各向度 通過率 寬轉長 表格 (不含 校+ 全體考生)
  socs.1<-gather(soc.1, key = 向度, value = 平均,-班級)
  socs.1[,3]<-round(x = socs.1[,3],digits = 2)
  # 將 各向度 通過率 寬轉長 表格 (含 校+ 全體考生)
  socs.2 <- gather(soc.all, key = 向度, value = 平均,-班級)
  socs.2[,3]<-round(x = socs.2[,3],digits = 2)
  #設定學校名 
  school<-  school.name
  
  # 重新設定 向度順序
  #Turn your 'treatment' column into a character vector
  socs$向度 <- as.character(socs$向度)
  #Then turn it back into an ordered factor
  socs$向度 <- factor(socs$向度, levels=unique(socs$向度))
  
  #Turn your 'treatment' column into a character vector
  socs.1$向度 <- as.character(socs.1$向度)
  #Then turn it back into an ordered factor
  socs.1$向度 <- factor(socs.1$向度, levels=unique(socs.1$向度))
  
  #Turn your 'treatment' column into a character vector
  soc.sall3$向度 <- as.character(soc.sall3$向度)
  #Then turn it back into an ordered factor
  soc.sall3$向度 <- factor(soc.sall3$向度, levels=unique(soc.sall3$向度))
  
  # 各班級 資料
  class.n <- nrow( soc.1 )
  
  if( plot.turn == TRUE){
    #1  該校 各班各向度答對率
    # 長條圖 數值設定
    g1.text    = if (cn<= 3 & subscale <= 10) {g1.text = 8} else{if(cn<= 5 & subscale <= 14){g1.text = 5}
      else{if(cn<= 10 & subscale <= 14){g1.text = 3}
        else{g1.text = 2 }}}
    g1.strip.text.x  = if (subscale <= 5) {g1.strip.text.x = 30} else{if(nl <=5){g1.strip.text.x = 26}else{g1.strip.text.x = 13}}
    
    g1.axis.text.x   = if (cn<=5) {g1.axis.text.x = 20} else{if(cn<=10){g1.axis.text.x = 10}else{g1.axis.text.x = 6}}
    g1.width         = if (cn >=10){ g1.width = 48} else{ g1.width = 40}
    g1.height        = if (cn >=10){ g1.height= 24} else{ g1.height= 20}
    g1.plot.title    = 40
    g1.axis.title.xy = 26
    g1.legend.title  = 36
    g1.legend.text   = 26
    g1.nrow          = if (subscale >=6){ g1.nrow = 2} else{ g1.nrow = 1}
    ################
    #用facet()，分別各畫一張各向度的長條圖
    p1<-ggplot(data=socs.1, aes(x=班級, y=平均,fill=班級 )) +
      geom_bar(colour="black",stat="identity", position=position_dodge(), size=.3 )  +
      geom_text(data=socs.1,aes(label=平均), vjust=1.6, color="white",
                position = position_dodge(0.9), size = g1.text  )+
      geom_hline(data=soc.sall3,aes(lty=群體,yintercept=平均),  size=1) +
      facet_wrap(~向度,nrow=g1.nrow)+
      #   facet_grid(.~向度)+
      guides(    color = guide_colorbar(order = 0),
                 fill = guide_legend(order = 1)  )+
      labs(title = paste( school, "各向度答對率", sep = " "),y = "答\n對\n率",family = "A")+
      theme( axis.text.x  = element_text(angle = 20, vjust = 0.5,size= g1.axis.text.x ,family = "A"),
             axis.text.y  = element_text(angle = 0, vjust = 0.5,size=30,family = "A"),
             strip.text.x = element_text(size = g1.strip.text.x ,family = "A"),
             plot.title   = element_text(hjust = 0.5,size= g1.plot.title ,family = "A"),
             axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g1.axis.title.xy,family = "A"),
             axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g1.axis.title.xy,family = "A"),
             legend.title = element_text(colour="black", size=g1.legend.title, face="bold",family = "A"),
             legend.text  = element_text( size = g1.legend.text,family = "A"))
    ## 匯入圖片
    print(p1) # plot needs to be showing
    insertPlot(wb, sheet = 2, startRow = (class.n+8), startCol = 2, width = g1.width, height = g1.height, fileType = "png", units = "cm")
    
    ############################################################################
    #2 各班各向度成績(校) 
    # 長條圖 數值設定
    g2.text          = if(cn<= 4){g2.text = 12}else{if(cn<=10){g2.text = 8}else{g2.text = 4}}
    
    g2.axis.text.x   = if(cn<= 5){g2.axis.text.x = 26}else{g2.axis.text.x = 16}
    g2.width         = if (cn+subscale >=20){ g1.width = 26} else{ g1.width = 22}
    g2.height        = if (cn+subscale >=20){ g1.height= 26} else{ g1.height= 22}
    g2.plot.title    = 22
    g2.axis.title.xy = 26
    g2.legend.title  = 36
    g2.legend.text   = 26
    
    ########################
    # cn 班級數量
    for(i in 1:cn) {
      soc.sub<-soc.1[,c(1,i+1)]
      soc.sub[,2]<-round(x = soc.sub[,2],digits = 2)
      soc.sub.names<- names(soc.1[-1])
      soc.sub.names<-soc.sub.names[i]
      
      # 畫圖 + 線
      # 問題 Y軸 名稱 .data[[soc.sub.names]]   
      p2 <-ggplot(data=soc.sub, aes_string(x="班級", y= soc.sub.names  ,fill="班級" )) +
        geom_bar(colour="black",stat="identity", position=position_dodge(),
                 size=.3)+
        # "群體"
        geom_hline(data=soc.sall2,aes_string(lty="群體",yintercept=soc.sub.names ),  size=1) +
        #geom_text(data=soc.sall2,aes(label = 群體), position = position_dodge(0.9))+
        geom_text(data=soc.sub,aes_string(label=soc.sub.names ), vjust=1.6, color="white",
                  position = position_dodge(0.9), size=g2.text)+
        guides(    color = guide_colorbar(order = 0),
                   fill = guide_legend(order = 1)  )+
        theme( axis.text.x = element_text(angle = 20, vjust = 0.5,size=g2.axis.text.x,family = "A"),
               axis.text.y = element_text(angle = 0, vjust = 0.5,size=20,family = "A"),
               plot.title = element_text(hjust = 0.5,size=g2.plot.title ,family = "A"),
               axis.title.x =element_text(hjust = 0.5,color="black",face="bold",size=g2.axis.title.xy,family = "A"),
               axis.title.y =element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g2.axis.title.xy,family = "A"),
               legend.title = element_text(colour="black", size=g2.legend.title, face="bold",family = "A"),
               legend.text = element_text( size = g2.legend.text,family = "A"))+
        labs(title = paste( school, soc.sub.names, sep = " "),y = "答\n對\n率",family = "A")
      ## 匯入圖片
      print(p2) # plot needs to be showing
      insertPlot(wb, sheet = 2, startRow = (class.n+50+44*(i-1)), startCol = 2, width = g2.width, height = g2.height, fileType = "png", units = "cm")
    }
    
    #####################################
    #3 班級成績
    # 長條圖 數值設定
    
    g3.text          = if(nl<=5 & subscale<=5){g3.text = 14}else{g3.text = 10}
    g3.strip.text.x  = if(nl<=7){g3.strip.text.x = 28}else{g3.strip.text.x = 12}
    
    g3.axis.text.x   = 34
    g3.width         = 40
    g3.height        = 20
    g3.plot.title    = 30
    g3.axis.title.xy = 28
    g3.legend.title  = 36
    g3.legend.text   = 26
    g3.nrow          = if (subscale >=10){ g1.nrow = 2} else{ g1.nrow = 1}
    ################
    
    
    for(i in 1:cn) {
      
      # 篩選 指定班級
      soc.classroom <- soc.1[i,]
      # 隱藏 校跟全體成績合併  
      # soc.classroom <- rbind(soc.classroom,soc.sall1)
      soc.classroom<-gather(soc.classroom, key = 向度, value = 平均,-班級)
      soc.classroom[,3]<-round(x = soc.classroom[,3],digits = 2)
      # 顯示 目前 班級
      classromm<- soc.1$班級[i]
      # 重新 factor 固定向度排序
      soc.classroom$向度<- as.character(soc.classroom$向度)
      soc.classroom$向度<- factor(soc.classroom$向度,levels=unique(socs.1$向度))
      # 計算 班級 數量
      classromm.n<-soc.s %>% filter(班級 == classromm) %>% nrow()
      
      p3<-ggplot(data=soc.classroom, aes(x=班級, y=平均,fill=班級 )) +
        geom_bar(colour="black",stat="identity", position=position_dodge(),
                 size=.3)+
        geom_text(data=soc.classroom,aes(label=平均), vjust=1.5, color="white",
                  position = position_dodge(0.9), size= g3.text )+
        geom_hline(data=soc.sall3,aes(lty=群體,yintercept=平均),  size=1) +
        facet_wrap(~向度,nrow = g3.nrow)+
        guides(    color = guide_colorbar(order = 0),
                   fill = guide_legend(order = 1)  )+
        theme( axis.text.x = element_text(angle = 0, vjust = 0.5,size=g3.axis.text.x,family = "A"),
               axis.text.y = element_text(angle = 0, vjust = 0.5,size=20,family = "A"),
               strip.text.x = element_text(size = g3.strip.text.x,family = "A"),
               plot.title = element_text(hjust = 0.5,size=g3.plot.title,family = "A"),
               axis.title.x =element_text(hjust = 0.5,color="black",face="bold",size=g3.axis.title.xy,family = "A"),
               axis.title.y =element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g3.axis.title.xy,family = "A"),
               legend.title = element_text(colour="black", size=g3.legend.title, face="bold",family = "A"),
               legend.text = element_text( size = g3.legend.text,family = "A"))+
        labs(title = paste( school, classromm ,"各向度答對率", sep = " "),y = "答\n對\n率",family = "A")
      
      ## 匯入圖片
      print(p3) # plot needs to be showing
      insertPlot(wb, sheet = (i+2) , startRow = (classromm.n+5), startCol = 2, width = g3.width, height = g3.height, fileType = "png", units = "cm")
    }
    ################################################################
    #########################################################
    #4 個人成績(班)
    # 長條圖 數值設定
    g4.text          =  5  
    g4.axis.text.x   = 13
    g4.width         = 40
    g4.height        = 20
    g4.plot.title    = 36
    g4.axis.title.xy = 30
    g4.legend.title  = 36
    g4.legend.text   = 26
    ########################
    # 先讀取 幾個 班級
    
    for(i in 1:cn) {
      for(j in 1:subscale){
        #指定 哪一班
        classromm<- soc.1$班級[i]
        sud.class<-soc.s %>% filter(班級 == classromm )
        # 班級人數
        classromm.n <- nrow(sud.class)
        # 指定 那一向度 ， 並 四捨五入至 小數第二位
        sud.class[,j+3] <- round(x = sud.class[,j+3] ,digits = 2)
        
        #      soc.sub<-soc.1[,c(j+1)]
        #      soc.sub[,2]<-round(x = soc.sub[,2],digits = 2)
        soc.sub.names<- names(soc.s[,-1:-3])
        soc.sub.names<-soc.sub.names[j]
        
        p4<-ggplot(data=sud.class, aes_string(x="姓名", y=soc.sub.names )) +
          geom_bar(fill="royalblue",stat="identity", position=position_dodge(),
                   size=.3)+
          geom_text(data=sud.class,aes_string(label=soc.sub.names), vjust=1.6, color="white",
                    position = position_dodge(0.9), size=g4.text)+
          geom_hline(data=soc.sall2,aes_string(lty="群體",yintercept=soc.sub.names),  size=1) +
          guides(    color = guide_colorbar(order = 0),
                     fill = guide_legend(order = 1)  )+
          theme( axis.text.x  = element_text(angle = 20, vjust = 0.5,size=g4.axis.text.x,family = "A"),
                 axis.text.y  = element_text(angle = 0, vjust = 0.5,size=30,family = "A"),
                 plot.title   = element_text(hjust = 0.5,size=g4.plot.title,family = "A"),
                 axis.title.x = element_text(hjust = 0.5,color="black",face="bold",size=g4.axis.title.xy,family = "A"),
                 axis.title.y = element_text(hjust = 0.5,color="black",angle=0, vjust = 0.5,face="bold",size=g4.axis.title.xy,family = "A"),
                 legend.title = element_text(colour="black", size=g4.legend.title, face="bold",family = "A"),
                 legend.text  = element_text( size = g4.legend.text,family = "A"))+
          labs(title = paste( school, classromm,soc.sub.names, sep = " "),y = "答\n對\n率",family = "A")
        ## 匯入圖片
        print(p4) # plot needs to be showing
        insertPlot(wb, sheet = (i+2) , startRow = (classromm.n+5+42*(j)), startCol = 2, width = g4.width, height = g4.height, fileType = "png", units = "cm")
      }
    }
  }
  
  ## 儲存新檔案
  saveWorkbook(wb, paste( "C:/Users/user/Desktop/",school,"109年學力預試報告",".xlsx",  sep=""), overwrite = TRUE)
  
}


##################################
