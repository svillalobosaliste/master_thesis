library(openxlsx)

setwd('//cbsp.nl/Productie/secundair/MPOnderzoek/Werk/Combineren/Projecten/Comb_Prob_Nonprob/code_revision2')

#FILES NEEDED: cond_eq.rds and cond_uneq.rds
cond_equal <- readRDS(file="cond_eq.rds")
cond_unequal <- readRDS(file="cond_uneq.rds")


#Outputs of this file are found on the folder "Outputs" -> "Bias"

cond_eq <- cond_equal[ , c("k","c","n.pk","n.npk","age.step","bias.nps","bias.comb","bias.EH","bias.IPF")]
cond_uneq <- cond_unequal[ , c("k","c","n.pk","n.npk","age.step","bias.nps","bias.comb","bias.EH","bias.IPF")]


##################################################
############## SPLIT PER SELECTIVITY #############
##################################################
bias_eq_05 <- subset(cond_eq, age.step == 0.5)
bias_eq_10 <- subset(cond_eq, age.step == 1)
bias_eq_15 <- subset(cond_eq, age.step == 1.5)
bias_uneq_05 <- subset(cond_uneq, age.step == 0.5)
bias_uneq_10 <- subset(cond_uneq, age.step == 1)
bias_uneq_15 <- subset(cond_uneq, age.step == 1.5)


##########################################################
#########COMPUTE DIFFERENCE OF BIAS NPS-COMB #############
##########################################################
#Difference of estimators of non-probability sample - combined estimator

bias_eq_05$diff <- (bias_eq_05$bias.nps - bias_eq_05$bias.comb)
bias_eq_10$diff <- (bias_eq_10$bias.nps - bias_eq_10$bias.comb)
bias_eq_15$diff <- (bias_eq_15$bias.nps - bias_eq_15$bias.comb)

bias_uneq_05$diff <- (bias_uneq_05$bias.nps - bias_uneq_05$bias.comb)
bias_uneq_10$diff <- (bias_uneq_10$bias.nps - bias_uneq_10$bias.comb)
bias_uneq_15$diff <- (bias_uneq_15$bias.nps - bias_uneq_15$bias.comb)

#a negative result means that the non-probability sample has a lower bias. 
#This should be marked in a Bold format in the final tables

bias_eq_05$diff.EH <- (bias_eq_05$bias.EH - bias_eq_05$bias.comb)
bias_eq_10$diff.EH <- (bias_eq_10$bias.EH - bias_eq_10$bias.comb)
bias_eq_15$diff.EH <- (bias_eq_15$bias.EH - bias_eq_15$bias.comb)

bias_uneq_05$diff.EH <- (bias_uneq_05$bias.EH - bias_uneq_05$bias.comb)
bias_uneq_10$diff.EH <- (bias_uneq_10$bias.EH - bias_uneq_10$bias.comb)
bias_uneq_15$diff.EH <- (bias_uneq_15$bias.EH - bias_uneq_15$bias.comb)

bias_eq_05$diff.IPF <- (bias_eq_05$bias.IPF - bias_eq_05$bias.comb)
bias_eq_10$diff.IPF <- (bias_eq_10$bias.IPF - bias_eq_10$bias.comb)
bias_eq_15$diff.IPF <- (bias_eq_15$bias.IPF - bias_eq_15$bias.comb)

bias_uneq_05$diff.IPF <- (bias_uneq_05$bias.IPF - bias_uneq_05$bias.comb)
bias_uneq_10$diff.IPF <- (bias_uneq_10$bias.IPF - bias_uneq_10$bias.comb)
bias_uneq_15$diff.IPF <- (bias_uneq_15$bias.IPF - bias_uneq_15$bias.comb)


##################################################################### 
################## FINAL TABLE DIFFERENCE OF BIAS ###################  
##################################################################### 

#With this function the column armse.comb is read and creates a table with armse values
bias.table <- function(data, kolom){
  
  K_vals <- sort(unique(data$k))
  C_vals <- sort(unique(data$c))
  p_vals <- sort(unique(data$n.pk))
  np_vals <- sort(unique(data$n.npk))
  
  res <- matrix(0, length(K_vals) * length(p_vals), length(C_vals) * length(np_vals))
  row.names(res) <- paste0(rep(paste0('K=', K_vals), each = length(p_vals)),
                           ', ',
                           rep(paste0('n.pk=', p_vals), times = length(K_vals)))
  colnames(res) <- paste0(rep(paste0('C=', C_vals), each = length(np_vals)),
                          ', ',
                          rep(paste0('n.npk=', np_vals), times = length(C_vals)))
  
  for (i in 1:length(C_vals)) {
    for (j in 1:length(np_vals)) {
      temp <- subset(data, c == C_vals[i] & n.npk == np_vals[j])
      res[ ,(i-1)*length(np_vals) + j] <- temp[order(temp$k, temp$n.pk), kolom]
    }
  }
  
  return(res)
}

b_diff_eq_05 <- bias.table(bias_eq_05, 'diff')
b_diff_eq_10 <- bias.table(bias_eq_10, 'diff')
b_diff_eq_15 <- bias.table(bias_eq_15, 'diff')

b_diff_uneq_05 <- bias.table(bias_uneq_05, 'diff')
b_diff_uneq_10 <- bias.table(bias_uneq_10, 'diff')
b_diff_uneq_15 <- bias.table(bias_uneq_15, 'diff')

writeExcel_layout(obj_data = b_diff_eq_05, name = 'bias_diff_eq_05_layout',
                  data = bias_eq_05, conditional = FALSE)
writeExcel_layout(obj_data = b_diff_eq_10, name = 'bias_diff_eq_10_layout',
                  data = bias_eq_10, conditional = FALSE)
writeExcel_layout(obj_data = b_diff_eq_15, name = 'bias_diff_eq_15_layout',
                  data = bias_eq_15, conditional = FALSE)
writeExcel_layout(obj_data = b_diff_uneq_05, name = 'bias_diff_uneq_05_layout',
                  data = bias_uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = b_diff_uneq_10, name = 'bias_diff_uneq_10_layout',
                  data = bias_uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = b_diff_uneq_15, name = 'bias_diff_uneq_15_layout',
                  data = bias_uneq_15, conditional = FALSE)



###################################################################
########COMPUTE PROPORTION OF COMBINED WITH LOWER BIAS#############  
###################################################################

#This function creates a new variable "prop" in which a 1 is computed if 
#the bias of the combined estimator is lower or equal to the estimator from the 
#non-probability sample, and a 0 if it is higher

prop.bias <- function (data){
  
  data$prop <- ifelse(data$diff >= 0, 1, 0)
  data$prop <- factor(data$prop, levels = c(0,1))
  
  return(data)
}

bias_eq_05 <- prop.bias(bias_eq_05)
bias_eq_10 <- prop.bias(bias_eq_10)
bias_eq_15 <- prop.bias(bias_eq_15)

bias_uneq_05 <- prop.bias(bias_uneq_05)
bias_uneq_10 <- prop.bias(bias_uneq_10)
bias_uneq_15 <- prop.bias(bias_uneq_15)

x1 <- table(bias_eq_05$prop)
x2 <- table(bias_eq_10$prop)
x3 <- table(bias_eq_15$prop)
x4 <- table(bias_uneq_05$prop)
x5 <- table(bias_uneq_10$prop)
x6 <- table(bias_uneq_15$prop)

proportion_bias <- rbind(x1,x2,x3,x4,x5,x6)

writeExcel(obj = proportion_bias, name = 'proportion_bias')


####
## Comparison to EH and IPF estimators

bias.less.EHandIPF <- matrix(0,6,3)
bias.less.EHandIPF[1,1] <- sum(bias_eq_05$diff.EH >= 0 | is.na(bias_eq_05$diff.EH))
bias.less.EHandIPF[1,2] <- sum(bias_eq_05$diff.IPF >= 0)
bias.less.EHandIPF[1,3] <- sum((bias_eq_05$diff.EH >= 0 | is.na(bias_eq_05$diff.EH)) & bias_eq_05$diff.IPF >= 0)

bias.less.EHandIPF[2,1] <- sum(bias_uneq_05$diff.EH >= 0 | is.na(bias_uneq_05$diff.EH))
bias.less.EHandIPF[2,2] <- sum(bias_uneq_05$diff.IPF >= 0)
bias.less.EHandIPF[2,3] <- sum((bias_uneq_05$diff.EH >= 0 | is.na(bias_uneq_05$diff.EH)) & bias_uneq_05$diff.IPF >= 0)

bias.less.EHandIPF[3,1] <- sum(bias_eq_10$diff.EH >= 0 | is.na(bias_eq_10$diff.EH))
bias.less.EHandIPF[3,2] <- sum(bias_eq_10$diff.IPF >= 0)
bias.less.EHandIPF[3,3] <- sum((bias_eq_10$diff.EH >= 0 | is.na(bias_eq_10$diff.EH)) & bias_eq_10$diff.IPF >= 0)

bias.less.EHandIPF[4,1] <- sum(bias_uneq_10$diff.EH >= 0 | is.na(bias_uneq_10$diff.EH))
bias.less.EHandIPF[4,2] <- sum(bias_uneq_10$diff.IPF >= 0)
bias.less.EHandIPF[4,3] <- sum((bias_uneq_10$diff.EH >= 0 | is.na(bias_uneq_10$diff.EH)) & bias_uneq_10$diff.IPF >= 0)

bias.less.EHandIPF[5,1] <- sum(bias_eq_15$diff.EH >= 0 | is.na(bias_eq_15$diff.EH))
bias.less.EHandIPF[5,2] <- sum(bias_eq_15$diff.IPF >= 0)
bias.less.EHandIPF[5,3] <- sum((bias_eq_15$diff.EH >= 0 | is.na(bias_eq_15$diff.EH)) & bias_eq_15$diff.IPF >= 0)

bias.less.EHandIPF[6,1] <- sum(bias_uneq_15$diff.EH >= 0 | is.na(bias_uneq_15$diff.EH))
bias.less.EHandIPF[6,2] <- sum(bias_uneq_15$diff.IPF >= 0)
bias.less.EHandIPF[6,3] <- sum((bias_uneq_15$diff.EH >= 0 | is.na(bias_uneq_15$diff.EH)) & bias_uneq_15$diff.IPF >= 0)

colnames(bias.less.EHandIPF)<-c("less than EH","less than IPF","less than EH and IPF")
rownames(bias.less.EHandIPF)<-c("eq_05","uneq_05",
                                "eq_10","uneq_10",
                                "eq_15","uneq_15")
bias.less.EHandIPF

bias.less.EHandIPF.1 <- (bias.less.EHandIPF*100/256)
bias.less.EHandIPF.1 <- round(bias.less.EHandIPF.1, digits=3)
bias.less.EHandIPF.1

writeExcel(obj = bias.less.EHandIPF.1, name = 'summary_bias_EHandIPF')



##################################################################### 
##################### BIAS OF COMBINED ESTIMATOR ####################  
##################################################################### 

# Now also make a table of the bias of the combined estimator

b_comb_eq_05 <- bias.table(bias_eq_05, 'bias.comb')
b_comb_eq_10 <- bias.table(bias_eq_10, 'bias.comb')
b_comb_eq_15 <- bias.table(bias_eq_15, 'bias.comb')

b_comb_uneq_05 <- bias.table(bias_uneq_05, 'bias.comb')
b_comb_uneq_10 <- bias.table(bias_uneq_10, 'bias.comb')
b_comb_uneq_15 <- bias.table(bias_uneq_15, 'bias.comb')

writeExcel_layout(obj_data = b_comb_eq_05, name = 'bias_comb_eq_05_layout',
                  data = bias_eq_05, conditional = FALSE)
writeExcel_layout(obj_data = b_comb_eq_10, name = 'bias_comb_eq_10_layout',
                  data = bias_eq_10, conditional = FALSE)
writeExcel_layout(obj_data = b_comb_eq_15, name = 'bias_comb_eq_15_layout',
                  data = bias_eq_15, conditional = FALSE)
writeExcel_layout(obj_data = b_comb_uneq_05, name = 'bias_comb_uneq_05_layout',
                  data = bias_uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = b_comb_uneq_10, name = 'bias_comb_uneq_10_layout',
                  data = bias_uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = b_comb_uneq_15, name = 'bias_comb_uneq_15_layout',
                  data = bias_uneq_15, conditional = FALSE)

