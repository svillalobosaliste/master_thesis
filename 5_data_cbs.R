library(openxlsx)

# function to write output as an Excel file with openxlsx
writeExcel <- function(obj, name) {
  # obj: name of R object to be written as output
  # name: intended name of Excel file (without .xlsx extension)
  wb <- createWorkbook()
  addWorksheet(wb, name)
  writeData(wb, name, obj,
            xy = c(1, 1), rowNames = TRUE)
  setColWidths(wb, name, 1:(ncol(obj)+1), widths = "auto")
  saveWorkbook(wb, file = paste0('Outputs/', name, '.xlsx'),
               overwrite = TRUE)
}

setwd('//cbsp.nl/Productie/secundair/MPOnderzoek/Werk/Combineren/Projecten/Comb_Prob_Nonprob/')

# Require data to CBS

#################################################################################
########### COMPUTING CBS ESTIMATES OF SELECTED CITIES : 3 CATEGORIES ########### 
#################################################################################
#Using the estimated from CBS Educational Attainment File
#with 18 educational levels, to recode it into 3 categories
#and select the 11 cities that will be use

#This CBS estimation includes both probability and non-probability sample
load('estimcbs.Rdata')

est$first<-est$N.1111+est$N.1112+est$N.1211+est$N.1212+est$N.1213+est$N.1221+est$N.1222
est$second<-est$N.2111+est$N.2112+est$N.2121+est$N.2131+est$N.2132
est$third<-est$N.3111+est$N.3112+est$N.3113+est$N.3211+est$N.3212+est$N.3213

cit1<-est[which(est$GEMJJJJ==c("0363")),]
cit2<-est[which(est$GEMJJJJ==c("0362")),]
cit3<-est[which(est$GEMJJJJ==c("1931")),]
cit4<-est[which(est$GEMJJJJ==c("0420")),]
cit5<-est[which(est$GEMJJJJ==c("1525")),]
cit6<-est[which(est$GEMJJJJ==c("0757")),]
cit7<-est[which(est$GEMJJJJ==c("0629")),]
cit8<-est[which(est$GEMJJJJ==c("1714")),]
cit9<-est[which(est$GEMJJJJ==c("0668")),]
cit10<-est[which(est$GEMJJJJ==c("0437")),]
cit11<-est[which(est$GEMJJJJ==c("0093")),]

cit<-rbind(cit1,cit2,cit3,cit4,cit5,cit6,cit7,cit8,cit9,cit10,cit11)
cities_3<-cbind(cit$firs,cit$second,cit$third)
cbs3<-prop.table(cities_3,1)

writeExcel(obj = round(cbs3,digits=2),
           name = 'cbs3')


###############################################################################
########### COMPUTING PROB-NONPROB ESTIMATES : 3 CATEGORIES ###################
###############################################################################
#Data from Labour Force Survey 2016 - This is the probability sample
#Select cities

EduProps_3 <- read.csv2('EduProps_3.csv')
Nobs_prob <- read.xlsx('NumObs_municipality.xlsx')

prob3<-EduProps_3
p1<-prob3[which(prob3$Municipality==c("363")),] #11051
p2<-prob3[which(prob3$Municipality==c("362")),] #1365
p3<-prob3[which(prob3$Municipality==c("1931")),] #1128
p4<-prob3[which(prob3$Municipality==c("420")),] # 845  
p5<-prob3[which(prob3$Municipality==c("1525")),] #684
p6<-prob3[which(prob3$Municipality==c("757")),] #683
p7<-prob3[which(prob3$Municipality==c("629")),] #336
p8<-prob3[which(prob3$Municipality==c("1714")),] #448
p9<-prob3[which(prob3$Municipality==c("668")),] #329
p10<-prob3[which(prob3$Municipality==c("437")),] #264
p11<-prob3[which(prob3$Municipality==c("93")),] #60

11051+1365+1128+845+684+683+336+448+329+264+60 #17193

prob3<-rbind(p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11)
prob.n <- sapply(c("363","362","1931","420","1525","757","629","1714","668","437","93"),
                 function(m) Nobs_prob$n[which(Nobs_prob$Municipality == m)])
prob3[[1]]<-NULL
prob3<-data.matrix(prob3)

writeExcel(obj = round(prob3,digits=2),
           name = 'prob3')


#####nonprob#####
#Using the observations from the non-probability sample of the Educational Attainment File

load('nonprob_table.Rdata')

x1<-as.data.frame(nonprob.table)

nonprob1<-x1[which(x1$Var1=="0363"),]
nonprob2<-x1[which(x1$Var1=="0362"),]
nonprob3<-x1[which(x1$Var1=="1931"),]
nonprob4<-x1[which(x1$Var1=="0420"),]
nonprob5<-x1[which(x1$Var1=="1525"),]
nonprob6<-x1[which(x1$Var1=="0757"),]
nonprob7<-x1[which(x1$Var1=="0629"),]
nonprob8<-x1[which(x1$Var1=="1714"),]
nonprob9<-x1[which(x1$Var1=="0668"),]
nonprob10<-x1[which(x1$Var1=="0437"),]
nonprob11<-x1[which(x1$Var1=="0093"),]

rr1<-t(nonprob1$Freq)
rr2<-t(nonprob2$Freq)
rr3<-t(nonprob3$Freq)
rr4<-t(nonprob4$Freq)
rr5<-t(nonprob5$Freq)
rr6<-t(nonprob6$Freq)
rr7<-t(nonprob7$Freq)
rr8<-t(nonprob8$Freq)
rr9<-t(nonprob9$Freq)
rr10<-t(nonprob10$Freq)
rr11<-t(nonprob11$Freq)

np<-as.data.frame(rbind(rr1,rr2,rr3,rr4,rr5,rr6,rr7,rr8,rr9,rr10,rr11))

np$first<-np$V1+np$V2+np$V3+np$V4+np$V5+np$V6+np$V7
np$second<-np$V8+np$V9+np$V10+np$V11+np$V12
np$third<-np$V13+np$V14+np$V15+np$V16+np$V17+np$V18

nonprob_est_freq<-cbind(np$first,np$second,np$third)
nonprob_est_prop<-prop.table(nonprob_est_freq,1)

nonprob.n<-rowSums(nonprob_est_freq)

writeExcel(obj = round(nonprob_est_prop,digits=2),
           name = 'nonprob3')


##############################################################################
#################### THE COMBINED ESTIMATOR: 3 CATEGORIES ####################
##############################################################################

#This function calculates the combined estimator - Model B for the combined samples
#of CBS data

c.est.function<-function(k,c,n.pk,n.npk,prob.table,nonprob.table,nonprob_est_prop,prob_est_prop){
  
  v.kc<-nonprob.table*(1-nonprob.table)
  bc<-colMeans(nonprob_est_prop-prob_est_prop)
  bc.mat<-matrix(bc,nrow=k,ncol=c,byrow=T)
  if (k > 1) {
    sigma2 <- sum((nonprob_est_prop - prob_est_prop - bc.mat)^2) / ((k-1)*c)
  } else {
    sigma2.k<-(n.pk/(n.pk-1))*((nonprob.table-prob.table)^2-(1-1/n.pk)*(bc^2)-(2/n.pk)*bc*nonprob.table)-v.kc*(1+n.npk/n.pk)/(n.npk-1)
    sigma2.k<-apply(sigma2.k,MARGIN = 1,FUN = mean)
    sigma2<-matrix(sigma2.k,nrow=k,ncol=c) 
  }
  wkc<-((n.npk-1)*(bc^2+sigma2)+v.kc)/((n.npk-1)*(1-1/n.pk)*(bc^2+sigma2)+(1+n.npk/n.pk)*v.kc+((n.npk-1)/n.pk)*bc*(2*nonprob.table-1))
  wkc[wkc<0]<-0
  wkc[wkc>1]<-1
  ccc<-wkc*prob.table+(1-wkc)*nonprob.table
  
  return(ccc)
  
}

#Specify sample size per domain of probability sample
n.pk1<-prob.n[1]
n.pk2<-prob.n[2]
n.pk3<-prob.n[3]
n.pk4<-prob.n[4]
n.pk5<-prob.n[5]
n.pk6<-prob.n[6]
n.pk7<-prob.n[7]
n.pk8<-prob.n[8]
n.pk9<-prob.n[9]
n.pk10<-prob.n[10]
n.pk11<-prob.n[11]

#Specify sample size per domain of non-probability sample
n.npk1<-nonprob.n[1]
n.npk2<-nonprob.n[2]
n.npk3<-nonprob.n[3]
n.npk4<-nonprob.n[4]
n.npk5<-nonprob.n[5]
n.npk6<-nonprob.n[6]
n.npk7<-nonprob.n[7]
n.npk8<-nonprob.n[8]
n.npk9<-nonprob.n[9]
n.npk10<-nonprob.n[10]
n.npk11<-nonprob.n[11]

#Specify proportion estimate from the probability sample
kp1<-t(as.matrix(prob3[1,]))
kp2<-t(as.matrix(prob3[2,]))
kp3<-t(as.matrix(prob3[3,]))
kp4<-t(as.matrix(prob3[4,]))
kp5<-t(as.matrix(prob3[5,]))
kp6<-t(as.matrix(prob3[6,]))
kp7<-t(as.matrix(prob3[7,]))
kp8<-t(as.matrix(prob3[8,]))
kp9<-t(as.matrix(prob3[9,]))
kp10<-t(as.matrix(prob3[10,]))
kp11<-t(as.matrix(prob3[11,]))

#Specify proportion estimate from the non-probability sample
knp1<-t(as.matrix(nonprob_est_prop[1,]))
knp2<-t(as.matrix(nonprob_est_prop[2,]))
knp3<-t(as.matrix(nonprob_est_prop[3,]))
knp4<-t(as.matrix(nonprob_est_prop[4,]))#
knp5<-t(as.matrix(nonprob_est_prop[5,]))
knp6<-t(as.matrix(nonprob_est_prop[6,]))
knp7<-t(as.matrix(nonprob_est_prop[7,]))
knp8<-t(as.matrix(nonprob_est_prop[8,]))
knp9<-t(as.matrix(nonprob_est_prop[9,]))
knp10<-t(as.matrix(nonprob_est_prop[10,]))
knp11<-t(as.matrix(nonprob_est_prop[11,]))

#Use the function to calculate combined estimator for the 11 cities
comb1<-c.est.function(11,3,n.pk1,n.npk1,kp1,knp1,nonprob_est_prop,prob3)
comb2<-c.est.function(11,3,n.pk2,n.npk2,kp2,knp2,nonprob_est_prop,prob3)
comb3<-c.est.function(11,3,n.pk3,n.npk3,kp3,knp3,nonprob_est_prop,prob3)
comb4<-c.est.function(11,3,n.pk4,n.npk4,kp4,knp4,nonprob_est_prop,prob3)
comb5<-c.est.function(11,3,n.pk5,n.npk5,kp5,knp5,nonprob_est_prop,prob3)
comb6<-c.est.function(11,3,n.pk6,n.npk6,kp6,knp6,nonprob_est_prop,prob3)
comb7<-c.est.function(11,3,n.pk7,n.npk7,kp7,knp7,nonprob_est_prop,prob3)
comb8<-c.est.function(11,3,n.pk8,n.npk8,kp8,knp8,nonprob_est_prop,prob3)
comb9<-c.est.function(11,3,n.pk9,n.npk9,kp9,knp9,nonprob_est_prop,prob3)
comb10<-c.est.function(11,3,n.pk10,n.npk10,kp10,knp10,nonprob_est_prop,prob3)
comb11<-c.est.function(11,3,n.pk11,n.npk11,kp11,knp11,nonprob_est_prop,prob3)

combined_estimator<-rbind(comb1,comb2,comb3,comb4,comb5,comb6,comb7,comb8,comb9,comb10,comb11) 

writeExcel(obj = round(combined_estimator,digits=2),
           name = 'combined_estimator')


####################################################
#################### SIMULATION ####################
####################################################

######prob####### 
#Taking two samples from CBS estimates 
prob.n                       #With same sample size
prob.smaller.n<-prob.n/10    #With smaller sample size

set.seed(4659)
samp.prob<-t(sapply(1:11,function(k)(rmultinom(n=1,size=prob.n[k],prob=cbs3[k,])))) 
samp.smaller<-t(sapply(1:11,function(k)(rmultinom(n=1,size=prob.smaller.n[k],prob=cbs3[k,]))))

prob_est_prop1<-prop.table(samp.prob,1)
prob_est_prop2<-prop.table(samp.smaller,1)

rowSums(samp.smaller)

##### Values for function #####

##Specify sample size per domain of sample with same sample size
n.pk1<-prob.n[1]  #is the same than rowSums(samp.prob)
n.pk2<-prob.n[2]
n.pk3<-prob.n[3]
n.pk4<-prob.n[4]
n.pk5<-prob.n[5]
n.pk6<-prob.n[6]
n.pk7<-prob.n[7]
n.pk8<-prob.n[8]
n.pk9<-prob.n[9]
n.pk10<-prob.n[10]
n.pk11<-prob.n[11]

##Specify sample size per domain of sample with smaller sample size
xn.pk1<-prob.smaller.n[1]
xn.pk2<-prob.smaller.n[2]
xn.pk3<-prob.smaller.n[3]
xn.pk4<-prob.smaller.n[4]
xn.pk5<-prob.smaller.n[5]
xn.pk6<-prob.smaller.n[6]
xn.pk7<-prob.smaller.n[7]
xn.pk8<-prob.smaller.n[8]
xn.pk9<-prob.smaller.n[9]
xn.pk10<-prob.smaller.n[10]
xn.pk11<-prob.smaller.n[11]

#Specify proportion estimate from the sample with same sample size
kp1<-t(as.matrix(prob_est_prop1[1,]))
kp2<-t(as.matrix(prob_est_prop1[2,]))
kp3<-t(as.matrix(prob_est_prop1[3,]))
kp4<-t(as.matrix(prob_est_prop1[4,]))
kp5<-t(as.matrix(prob_est_prop1[5,]))
kp6<-t(as.matrix(prob_est_prop1[6,]))
kp7<-t(as.matrix(prob_est_prop1[7,]))
kp8<-t(as.matrix(prob_est_prop1[8,]))
kp9<-t(as.matrix(prob_est_prop1[9,]))
kp10<-t(as.matrix(prob_est_prop1[10,]))
kp11<-t(as.matrix(prob_est_prop1[11,]))

#Specify proportion estimate from the sample with smaller sample size
xkp1<-t(as.matrix(prob_est_prop2[1,]))
xkp2<-t(as.matrix(prob_est_prop2[2,]))
xkp3<-t(as.matrix(prob_est_prop2[3,]))
xkp4<-t(as.matrix(prob_est_prop2[4,]))
xkp5<-t(as.matrix(prob_est_prop2[5,]))
xkp6<-t(as.matrix(prob_est_prop2[6,]))
xkp7<-t(as.matrix(prob_est_prop2[7,]))
xkp8<-t(as.matrix(prob_est_prop2[8,]))
xkp9<-t(as.matrix(prob_est_prop2[9,]))
xkp10<-t(as.matrix(prob_est_prop2[10,]))
xkp11<-t(as.matrix(prob_est_prop2[11,]))

#Use the function to calculate combined estimator for the 11 cities for the sample with same sample size
#non-probability sample values are the same as the first calculation of combined estimator
comb1<-c.est.function(11,3,n.pk1,n.npk1,kp1,knp1,nonprob_est_prop,prob_est_prop1)
comb2<-c.est.function(11,3,n.pk2,n.npk2,kp2,knp2,nonprob_est_prop,prob_est_prop1)
comb3<-c.est.function(11,3,n.pk3,n.npk3,kp3,knp3,nonprob_est_prop,prob_est_prop1)
comb4<-c.est.function(11,3,n.pk4,n.npk4,kp4,knp4,nonprob_est_prop,prob_est_prop1)
comb5<-c.est.function(11,3,n.pk5,n.npk5,kp5,knp5,nonprob_est_prop,prob_est_prop1)
comb6<-c.est.function(11,3,n.pk6,n.npk6,kp6,knp6,nonprob_est_prop,prob_est_prop1)
comb7<-c.est.function(11,3,n.pk7,n.npk7,kp7,knp7,nonprob_est_prop,prob_est_prop1)
comb8<-c.est.function(11,3,n.pk8,n.npk8,kp8,knp8,nonprob_est_prop,prob_est_prop1)
comb9<-c.est.function(11,3,n.pk9,n.npk9,kp9,knp9,nonprob_est_prop,prob_est_prop1)
comb10<-c.est.function(11,3,n.pk10,n.npk10,kp10,knp10,nonprob_est_prop,prob_est_prop1)
comb11<-c.est.function(11,3,n.pk11,n.npk11,kp11,knp11,nonprob_est_prop,prob_est_prop1)

#Use the function to calculate combined estimator for the 11 cities for the sample with smaller sample size
#non-probability sample values are the same as the first calculation of combined estimator
c1<-c.est.function(11,3,xn.pk1,n.npk1,xkp1,knp1,nonprob_est_prop,prob_est_prop2)
c2<-c.est.function(11,3,xn.pk2,n.npk2,xkp2,knp2,nonprob_est_prop,prob_est_prop2)
c3<-c.est.function(11,3,xn.pk3,n.npk3,xkp3,knp3,nonprob_est_prop,prob_est_prop2)
c4<-c.est.function(11,3,xn.pk4,n.npk4,xkp4,knp4,nonprob_est_prop,prob_est_prop2)
c5<-c.est.function(11,3,xn.pk5,n.npk5,xkp5,knp5,nonprob_est_prop,prob_est_prop2)
c6<-c.est.function(11,3,xn.pk6,n.npk6,xkp6,knp6,nonprob_est_prop,prob_est_prop2)
c7<-c.est.function(11,3,xn.pk7,n.npk7,xkp7,knp7,nonprob_est_prop,prob_est_prop2)
c8<-c.est.function(11,3,xn.pk8,n.npk8,xkp8,knp8,nonprob_est_prop,prob_est_prop2)
c9<-c.est.function(11,3,xn.pk9,n.npk9,xkp9,knp9,nonprob_est_prop,prob_est_prop2)
c10<-c.est.function(11,3,xn.pk10,n.npk10,xkp10,knp10,nonprob_est_prop,prob_est_prop2)
c11<-c.est.function(11,3,xn.pk11,n.npk11,xkp11,knp11,nonprob_est_prop,prob_est_prop2)

combined_big_sample<-rbind(comb1,comb2,comb3,comb4,comb5,comb6,comb7,comb8,comb9,comb10,comb11) 
combined_small_sample<-rbind(c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11) #sample size of probability sample divided by 10

writeExcel(obj = round(combined_big_sample,digits=2),
           name = 'combined_big_sample')
writeExcel(obj = round(combined_small_sample,digits=2),
           name = 'combined_small_sample')
