library(reticulate)
library(rio)
library(dplyr)
library(openxlsx)
library(readxl)
library(dummies)
library(data.table)
library(ggplot2)
library(tidyverse)
library(ggthemes)
# 第二種 設計 : 在 R 中 整理 資料 + 圖 ， 在將其存入新的 excel 中
# 參考 計分程式
#source("C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\00程式需求\\預試程式\\科目答案(預試程式版0612).R", encoding = "utf-8")
# 參考 繪圖程式
#source("C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\00程式需求\\預試程式\\預試成績繪圖0612(備份用).R", encoding = "utf-8")

# source("C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\00程式需求\\預試程式\\RtoEXCEL程式(0616).R", encoding = "utf-8")
# ## 匯入原始資料路徑  # choose.files()
# data.path = "C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\20200409-學力預試分析報告\\20200416-學力預試-姓名\\國小\\數學\\彰化-田中-五年-數學.xlsx"
# # 匯入 預設 全體通過率
# soc.all.in.one <- read_excel("C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\00程式需求\\預試程式\\數學答案.xlsx",sheet = "4all")
# ## 匯入 資料 與 答案
# data <- read_excel(data.path)
# key  <- read_excel("C:\\Users\\user\\Documents\\台中教育大學\\學力工作\\00程式需求\\預試程式\\數學答案.xlsx",sheet = "4A")

school.score.plot.cc<-function(data.path,data,soc.all.in.one,key,plot.turn) {
# 建立 workbook
wb <- createWorkbook()
## 取得 檔名
dn<-data.path %>% basename() %>% strsplit( "[,. ]") 
school.name<-dn[[1]][1]
# 固定 班級為類別
data$班級<- as.factor(data$班級) 


# 第一部份 : 計算分數
# 注意 題目的順序  必須正確(1.2.3.4....n)
#班級 必為類別
data$班級 <- as.factor(data$班級) 
#前 8 欄是學生基本資料
data.q<- data[,-1:-8] 
# 注意 答案的順序 必須正確(1.2.3.4....n)
key.a <- c(key$正確答案) 

#設定學校名 
school<-  school.name

# 根據 正確答案 核對出 二元作答反應
files<- length(c(data[[1]]))
df.1 <- data.frame()
df.2 <- data.frame()
for(i in 1:files) {
  df.1 <- ifelse( data.q[i,] == (key.a), 1, 0)
  df.2 <- rbind(df.2, df.1)  
}
# 將 NA部分處理為 0
df.2[is.na(df.2)  == TRUE]  <- 0 

總平均 <- rowSums(df.2)/ncol(df.2)
# 分出向度
key.dummy<- data.table(dummy(key$評量向度))
names(key.dummy)<-c(levels(factor(key$評量向度)))

# 計算向度總分平均
n.scales <- ncol(key.dummy)
sets <- apply(key.dummy, 2, function(XXX) data.table(df.2[, XXX == 1]))
names(sets) <- c(levels(factor(key$評量向度)))
# 計算向度平均值
list.1<-lapply(sets,function(x)  rowSums(x)/ncol(x))
d1<-do.call("cbind",list.1)

# 按照 題數 更改 向度名稱 ex: "代數(3)"
# 計算 向度 數量
n.scales <- ncol(key.dummy)
# 取得 向度 名稱
sub.names = colnames(d1)
# 用迴圈 將 向度名稱 更換
for(i in 1:n.scales) {
  sub.n = sum( key.dummy[[i]] ) 
  sub.names[i] <- paste( sub.names[i], "(",sub.n, ")", sep="")
}
# 取代 原有向度名稱
d2<-d1
colnames(d2) = sub.names
# 輸出計算結果 已 班級為單位
soc.1<-data.table(班級=data$班級,總平均,d2)
soc.2<-data.table(班級="校平均",總平均,d2)
#  soc.3<-data.table(全體通過率=c(key$`CTT難度(%)`),評量向度=c(key$評量向度))
# 計算 班級 為分類 的 平均(總平均+各向度)
soc1<-soc.1   %>%
  group_by(班級) %>%
  summarise_if(is.numeric,mean, na.rm = TRUE)
# 計算 校 為分類 的 平均(總平均+各向度)
soc2<-soc.2   %>%
  group_by(班級) %>%
  summarise_if(is.numeric,mean, na.rm = TRUE)
# 合併 各班成績 與 校成績 加 全體通過率
soc3<- rbind(soc1,soc2,soc.all.in.one)
# 合併 班級、座號、姓名 + 成績
soc.s<-data.table(學校 = data$學校,班級=data$班級,姓名 = data$姓名,總平均,d1)

# 加入註釋
note1<- data.table(班級="註1:全體通過率是指參與本次預試的全國所有學生平均通過率。")
note2<- data.table(班級="註2:標題刮號內數字表示各向度包含之題數。")
note3<- full_join( note1,note2, by = "班級")
soc3.note <- full_join(soc3, note3, by = "班級")
# 整合 匯出資料
out<-data.table(data[,1:8],總平均,d1)
# 建立分頁 & 匯入 對應 資料
addWorksheet(wb, sheetName = "個人成績(校)") 
addWorksheet(wb, sheetName = "各班成績(校)")
## 匯入 對應data
writeData(wb, sheet = "個人成績(校)", x = out)
writeData(wb, sheet = "各班成績(校)", x = soc3.note)

# 各班級 資料
class.n <- nrow( soc1 )
for(i in 1:class.n) {
  classromm <- soc1$班級[i]
  class.out <- out %>% filter(班級 == classromm)
  # 加入註釋
  note4<- data.table(卷別="註1:全體通過率是指參與本次預試的全國所有學生平均通過率。")
  class.out.note <- full_join(class.out, note4, by = "卷別")
  # paste( "個人&班級成績(",classromm,")" , sep="")
  addWorksheet(wb, sheetName = paste( "個人與班級成績","(",classromm ,")", sep="")) 
  writeData(wb, sheet = paste( "個人與班級成績","(",classromm ,")", sep=""), x = class.out.note)
}

##################################################################################################

# 第二部份 : 繪製圖片 & 匯入 對應  圖片
windowsFonts(A=windowsFont("標楷體"))
# 換回 原有向度名稱
colnames(soc1)<- data.table(班級=data$班級,總平均,d1) %>% colnames()
colnames(soc2)<- data.table(班級=data$班級,總平均,d1) %>% colnames()
colnames(soc3)<- data.table(班級=data$班級,總平均,d1) %>% colnames()
# 計算 班級數量
cn<-nrow(soc1)
# 計算 向度數量(含總平均) 扣除 班級欄位 1欄
subscale <- ncol(soc3)-1
# 計算 向度最長字數
nl<-data.table(d1) %>% names() %>% nchar() %>% max()
soc.all<- soc3
#排除 班級的成績
soc.sall1<-soc3[-1:-cn,]
soc.sall2<-soc.sall1 %>% 
  rename(
    群體 = 班級
  )
# 將 各向度 通過率 寬轉長 表格
soc.sall3<-gather(soc.sall2, key = 向度, value = 平均,-群體)
# 直接匯入
socs<-gather(soc3, key = 向度, value = 平均,-班級)
socs[,3]<-round(x = socs[,3],digits = 2)
# 將 各向度 通過率 寬轉長 表格 (不含 校+ 全體考生)
socs.1<-gather(soc1, key = 向度, value = 平均,-班級)
socs.1[,3]<-round(x = socs.1[,3],digits = 2)
# 將 各向度 通過率 寬轉長 表格 (含 校+ 全體考生)
socs.2 <- gather(soc3, key = 向度, value = 平均,-班級)
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
########################################################################
if( plot.turn == TRUE){
#1  該校 各班各向度答對率
# 長條圖 數值設定
g1.text          = if (cn<= 3 & subscale <= 10) {g1.text = 8} else{if(cn<= 5 & subscale <= 14){g1.text = 5}
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
  soc.sub<-soc1[,c(1,i+1)]
  soc.sub[,2]<-round(x = soc.sub[,2],digits = 2)
  soc.sub.names<- names(soc1[-1])
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
  soc.classroom <- soc1[i,]
  # 隱藏 校跟全體成績合併  
  # soc.classroom <- rbind(soc.classroom,soc.sall1)
  soc.classroom<-gather(soc.classroom, key = 向度, value = 平均,-班級)
  soc.classroom[,3]<-round(x = soc.classroom[,3],digits = 2)
  # 顯示 目前 班級
  classromm<- soc1$班級[i]
  # 重新 factor 固定向度排序
  soc.classroom$向度<- as.character(soc.classroom$向度)
  soc.classroom$向度<- factor(soc.classroom$向度,levels=unique(socs.1$向度))
  # 計算 班級 數量
  classromm.n<-soc.1 %>% filter(班級 == classromm) %>% nrow()
  
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
    classromm<- soc1$班級[i]
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




