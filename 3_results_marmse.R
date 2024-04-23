library(openxlsx)
library(ggplot2)

setwd('//cbsp.nl/Productie/secundair/MPOnderzoek/Werk/Combineren/Projecten/Comb_Prob_Nonprob/code_revision2')

#FILES NEEDED: cond_eq.rds and cond_uneq.rds
cond_equal <- readRDS(file="cond_eq.rds")
cond_unequal <- readRDS(file="cond_uneq.rds")

#Outputs of this file are found on the folder "Outputs" -> "MARMSE"

writeExcel(obj = cond_equal, name = 'cond_eq')
writeExcel(obj = cond_unequal, name = 'cond_uneq')


#############################

cond_equal$sel.IPF <- cond_equal$armse.IPF >= cond_equal$armse.comb
cond_unequal$sel.IPF <- cond_unequal$armse.IPF >= cond_unequal$armse.comb

table(cond_equal$sel.IPF, useNA = 'ifany')
table(cond_unequal$sel.IPF, useNA = 'ifany')

L_IPF_equal <- step(glm(sel.IPF ~ 1, family = binomial, data = cond_equal),
                    scope = list(lower = 'sel.IPF ~ 1',
                                 upper = 'sel.IPF ~ factor(k) + factor(c) + factor(n.pk) + factor(n.npk) + factor(age.step)'),
                    direction = 'forward')
summary(L_IPF_equal)

L_IPF_unequal <- step(glm(sel.IPF ~ 1, family = binomial, data = cond_unequal),
                      scope = list(lower = 'sel.IPF ~ 1',
                                   upper = 'sel.IPF ~ factor(k) + factor(c) + factor(n.pk) + factor(n.npk) + factor(age.step)'),
                      direction = 'forward')
summary(L_IPF_unequal)

dat <- rbind(cond_equal, cond_unequal)
dat$cat.type <- factor(c(rep('equal',nrow(cond_equal)),
                         rep('unequal', nrow(cond_unequal))))

L_IPF <- step(glm(sel.IPF ~ 1, family = binomial, data = dat),
              scope = list(lower = 'sel.IPF ~ 1',
                           upper = 'sel.IPF ~ factor(k) + factor(c) + factor(n.pk) + factor(n.npk) + factor(age.step) + cat.type'),
              direction = 'forward')
summary(L_IPF)


#############################

cond_eq <- cond_equal[ , c("k","c","n.pk","n.npk","age.step","armse.ps","armse.nps","armse.comb","armse.EH","armse.IPF")]
cond_uneq <- cond_unequal[ , c("k","c","n.pk","n.npk","age.step","armse.ps","armse.nps","armse.comb","armse.EH","armse.IPF")]


##################################################
### COUNTING HOW MANY TIMES COMBINED ESTIMATOR ###
#####   IS LOWER THAN: PROBABILITY SAMPLE,   #####
#####   NON-PROBABILITY SAMPLE AND BOTH      #####
##################################################

less_ps_eq <- subset(cond_eq, armse.comb < armse.ps)       # 761
less_ps_uneq <- subset(cond_uneq, armse.comb < armse.ps)   # 745

less_nps_eq <- subset(cond_eq, armse.comb < armse.nps)     # 689
less_nps_uneq <- subset(cond_uneq, armse.comb < armse.nps) # 689

less_both_eq <- subset(cond_eq, (armse.comb < armse.ps) & (armse.comb < armse.nps))     # 682
less_both_uneq <- subset(cond_uneq, (armse.comb < armse.ps) & (armse.comb < armse.nps)) # 666

##################################################
### COUNTING HOW MANY TIMES COMBINED ESTIMATOR ###
#####   IS LOWER THAN: ELLIOTT-HAVILAND,     #####
#####   IPF AND BOTH                         #####
##################################################

less_EH_eq <- subset(cond_eq, armse.comb < armse.EH | is.na(armse.EH))       # 755
less_EH_uneq <- subset(cond_uneq, armse.comb < armse.EH | is.na(armse.EH))   # 759

less_IPF_eq <- subset(cond_eq, armse.comb < armse.IPF)     # 396
less_IPF_uneq <- subset(cond_uneq, armse.comb < armse.IPF) # 407

less_EHandIPF_eq <- subset(cond_eq, (armse.comb < armse.EH | is.na(armse.EH)) & (armse.comb < armse.IPF))     # 383
less_EHandIPF_uneq <- subset(cond_uneq, (armse.comb < armse.EH | is.na(armse.EH)) & (armse.comb < armse.IPF)) # 398

#############################################################
### CREATING FUNCTION THAT STORAGES IF COMBINED ESTIMATOR ###
#######   IS LOWER AS A DUMMY VARIABLE YES=1 NO=0     #######
#############################################################

less.than.function <- function(data) {
  
  data$less_ps <- as.integer(data$armse.comb < data$armse.ps)
  data$less_nps <- as.integer(data$armse.comb < data$armse.nps)
  data$less_both <- data$less_ps * data$less_nps
  
  data$less_EH <- as.integer(data$armse.comb < data$armse.EH | is.na(data$armse.EH))
  data$less_IPF <- as.integer(data$armse.comb < data$armse.IPF)
  data$less_EHandIPF <- as.integer((data$armse.comb < data$armse.EH | is.na(data$armse.EH)) & data$armse.comb < data$armse.IPF)
  
  return(data)
  
}

armse.eq <- less.than.function(cond_eq)
armse.uneq <- less.than.function(cond_uneq)


################################################
#########  SPLITING PER SELECTIVITY    ######### 
################################################

eq_05 <- subset(armse.eq, age.step == 0.5)
eq_10 <- subset(armse.eq, age.step == 1)
eq_15 <- subset(armse.eq, age.step == 1.5)
uneq_05 <- subset(armse.uneq, age.step == 0.5)
uneq_10 <- subset(armse.uneq, age.step == 1)
uneq_15 <- subset(armse.uneq, age.step == 1.5)

saveRDS(eq_05, file = "eq_05.rds")
saveRDS(eq_10, file = "eq_10.rds")
saveRDS(eq_15, file = "eq_15.rds")
saveRDS(uneq_05, file = "uneq_05.rds")
saveRDS(uneq_10, file = "uneq_10.rds")
saveRDS(uneq_15, file = "uneq_15.rds")



##########################################################################
#########  COMPUTE PROPORTION OF TIMES COMBINED HAS LOWER MSE    ######### 
##########################################################################

sum.less <- matrix(0,6,3)
sum.less[1,1] <- sum(eq_05$less_ps, na.rm = TRUE)
sum.less[1,2] <- sum(eq_05$less_nps, na.rm = TRUE)
sum.less[1,3] <- sum(eq_05$less_both, na.rm = TRUE)

sum.less[2,1] <- sum(uneq_05$less_ps, na.rm = TRUE)
sum.less[2,2] <- sum(uneq_05$less_nps, na.rm = TRUE)
sum.less[2,3] <- sum(uneq_05$less_both, na.rm = TRUE)

sum.less[3,1] <- sum(eq_10$less_ps, na.rm = TRUE)
sum.less[3,2] <- sum(eq_10$less_nps, na.rm = TRUE)
sum.less[3,3] <- sum(eq_10$less_both, na.rm = TRUE)

sum.less[4,1] <- sum(uneq_10$less_ps, na.rm = TRUE)
sum.less[4,2] <- sum(uneq_10$less_nps, na.rm = TRUE)
sum.less[4,3] <- sum(uneq_10$less_both, na.rm = TRUE)

sum.less[5,1] <- sum(eq_15$less_ps, na.rm = TRUE)
sum.less[5,2] <- sum(eq_15$less_nps, na.rm = TRUE)
sum.less[5,3] <- sum(eq_15$less_both, na.rm = TRUE)

sum.less[6,1] <- sum(uneq_15$less_ps, na.rm = TRUE)
sum.less[6,2] <- sum(uneq_15$less_nps, na.rm = TRUE)
sum.less[6,3] <- sum(uneq_15$less_both, na.rm = TRUE)

colnames(sum.less)<-c("less than ps","less than nps","less than both")
rownames(sum.less)<-c("eq_05","uneq_05",
                      "eq_10","uneq_10",
                      "eq_15","uneq_15")
sum.less

sum.less.1 <- (sum.less*100/256)
sum.less.1 <- round(sum.less.1, digits=3)
sum.less.1

writeExcel(obj = sum.less.1, name = 'summary_marmse')


sum.less.EHandIPF <- matrix(0,6,3)
sum.less.EHandIPF[1,1] <- sum(eq_05$less_EH, na.rm = TRUE)
sum.less.EHandIPF[1,2] <- sum(eq_05$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[1,3] <- sum(eq_05$less_EHandIPF, na.rm = TRUE)

sum.less.EHandIPF[2,1] <- sum(uneq_05$less_EH, na.rm = TRUE)
sum.less.EHandIPF[2,2] <- sum(uneq_05$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[2,3] <- sum(uneq_05$less_EHandIPF, na.rm = TRUE)

sum.less.EHandIPF[3,1] <- sum(eq_10$less_EH, na.rm = TRUE)
sum.less.EHandIPF[3,2] <- sum(eq_10$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[3,3] <- sum(eq_10$less_EHandIPF, na.rm = TRUE)

sum.less.EHandIPF[4,1] <- sum(uneq_10$less_EH, na.rm = TRUE)
sum.less.EHandIPF[4,2] <- sum(uneq_10$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[4,3] <- sum(uneq_10$less_EHandIPF, na.rm = TRUE)

sum.less.EHandIPF[5,1] <- sum(eq_15$less_EH, na.rm = TRUE)
sum.less.EHandIPF[5,2] <- sum(eq_15$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[5,3] <- sum(eq_15$less_EHandIPF, na.rm = TRUE)

sum.less.EHandIPF[6,1] <- sum(uneq_15$less_EH, na.rm = TRUE)
sum.less.EHandIPF[6,2] <- sum(uneq_15$less_IPF, na.rm = TRUE)
sum.less.EHandIPF[6,3] <- sum(uneq_15$less_EHandIPF, na.rm = TRUE)

colnames(sum.less.EHandIPF)<-c("less than EH","less than IPF","less than EH and IPF")
rownames(sum.less.EHandIPF)<-c("eq_05","uneq_05",
                               "eq_10","uneq_10",
                               "eq_15","uneq_15")
sum.less.EHandIPF

sum.less.EHandIPF.1 <- (sum.less.EHandIPF*100/256)
sum.less.EHandIPF.1 <- round(sum.less.EHandIPF.1, digits=3)
sum.less.EHandIPF.1

writeExcel(obj = sum.less.EHandIPF.1, name = 'summary_marmse_EHandIPF')


#################################################
##########  CREATE TABLES MARMSE (8)   ##########
#################################################

#With this function the column armse.comb is read and creates a table with armse values
marmse.table <- function(data, kolom){
  
  K_vals <- sort(unique(data$k))
  C_vals <- sort(unique(data$c))
  p_vals <- sort(unique(data$n.pk))
  np_vals <- sort(unique(data$n.npk))
  
  marmse <- matrix(0, length(K_vals) * length(p_vals), length(C_vals) * length(np_vals))
  row.names(marmse) <- paste0(rep(paste0('K=', K_vals), each = length(p_vals)),
                              ', ',
                              rep(paste0('n.pk=', p_vals), times = length(K_vals)))
  colnames(marmse) <- paste0(rep(paste0('C=', C_vals), each = length(np_vals)),
                             ', ',
                             rep(paste0('n.npk=', np_vals), times = length(C_vals)))
  
  for (i in 1:length(C_vals)) {
    for (j in 1:length(np_vals)) {
      temp <- subset(data, c == C_vals[i] & n.npk == np_vals[j])
      marmse[ ,(i-1)*length(np_vals) + j] <- temp[order(temp$k, temp$n.pk), kolom]
    }
  }
  
  return(marmse)
}

marmse_eq_05 <- marmse.table(eq_05, 'armse.comb')
marmse_eq_10 <- marmse.table(eq_10, 'armse.comb')
marmse_eq_15 <- marmse.table(eq_15, 'armse.comb')

marmse_uneq_05 <- marmse.table(uneq_05, 'armse.comb')
marmse_uneq_10 <- marmse.table(uneq_10, 'armse.comb')
marmse_uneq_15 <- marmse.table(uneq_15, 'armse.comb')

writeExcel(obj = marmse_eq_05, name = 'marmse_eq_05')
writeExcel(obj = marmse_eq_10, name = 'marmse_eq_10')
writeExcel(obj = marmse_eq_15, name = 'marmse_eq_15')
writeExcel(obj = marmse_uneq_05, name = 'marmse_uneq_05')
writeExcel(obj = marmse_uneq_10, name = 'marmse_uneq_10')
writeExcel(obj = marmse_uneq_15, name = 'marmse_uneq_15')



#################################################
########  CREATE GROUPS TABLE MARMSE (8) ####### 
#################################################

# This calculation indicates if the value of the MARMSE of the combined estimator is lower 
#than the probability sample, lower than the non-probability sample, or lower than both.
# 2 = means that it is lower than the probability sample
# 3 = means that it is lower than the non-probability sample  
# 6 = means that it is lower than both 

eq_05$color <- eq_05$less_ps*2 + eq_05$less_nps*3 + eq_05$less_both*1
eq_10$color <- eq_10$less_ps*2 + eq_10$less_nps*3 + eq_10$less_both*1
eq_15$color <- eq_15$less_ps*2 + eq_15$less_nps*3 + eq_15$less_both*1

uneq_05$color <- uneq_05$less_ps*2 + uneq_05$less_nps*3 + uneq_05$less_both*1
uneq_10$color <- uneq_10$less_ps*2 + uneq_10$less_nps*3 + uneq_10$less_both*1
uneq_15$color <- uneq_15$less_ps*2 + uneq_15$less_nps*3 + uneq_15$less_both*1

ew <- marmse.table(eq_05, 'color')
em <- marmse.table(eq_10, 'color')
es <- marmse.table(eq_15, 'color')

uw <- marmse.table(uneq_05, 'color')
um <- marmse.table(uneq_10, 'color')
us <- marmse.table(uneq_15, 'color')

writeExcel(obj = ew, name = 'marmse_eq_05_color')
writeExcel(obj = em, name = 'marmse_eq_10_color')
writeExcel(obj = es, name = 'marmse_eq_15_color')
writeExcel(obj = uw, name = 'marmse_uneq_05_color')
writeExcel(obj = um, name = 'marmse_uneq_10_color')
writeExcel(obj = us, name = 'marmse_uneq_15_color')

#The MARMSE tables are dependent on the color tables on excel, and are colored according to these
#values as indicated in the report.

writeExcel_layout(obj_data = marmse_eq_05, obj_color = ew,
                  name = 'marmse_eq_05_layout', data = eq_05, conditional = TRUE)
writeExcel_layout(obj_data = marmse_eq_10, obj_color = em,
                  name = 'marmse_eq_10_layout', data = eq_10, conditional = TRUE)
writeExcel_layout(obj_data = marmse_eq_15, obj_color = es,
                  name = 'marmse_eq_15_layout', data = eq_15, conditional = TRUE)
writeExcel_layout(obj_data = marmse_uneq_05, obj_color = uw,
                  name = 'marmse_uneq_05_layout', data = uneq_05, conditional = TRUE)
writeExcel_layout(obj_data = marmse_uneq_10, obj_color = um,
                  name = 'marmse_uneq_10_layout', data = uneq_10, conditional = TRUE)
writeExcel_layout(obj_data = marmse_uneq_15, obj_color = us,
                  name = 'marmse_uneq_15_layout', data = uneq_15, conditional = TRUE)



###################
# Now do the same for the armse of the probability sample and the non-probability sample

ps_eq_05 <- marmse.table(eq_05, 'armse.ps')
ps_eq_10 <- marmse.table(eq_10, 'armse.ps')
ps_eq_15 <- marmse.table(eq_15, 'armse.ps')

ps_uneq_05 <- marmse.table(uneq_05, 'armse.ps')
ps_uneq_10 <- marmse.table(uneq_10, 'armse.ps')
ps_uneq_15 <- marmse.table(uneq_15, 'armse.ps')

writeExcel_layout(obj_data = ps_eq_05, name = 'ps_marmse_eq_05_layout',
                  data = eq_05, conditional = FALSE)
writeExcel_layout(obj_data = ps_eq_10, name = 'ps_marmse_eq_10_layout',
                  data = eq_10, conditional = FALSE)
writeExcel_layout(obj_data = ps_eq_15, name = 'ps_marmse_eq_15_layout',
                  data = eq_15, conditional = FALSE)
writeExcel_layout(obj_data = ps_uneq_05, name = 'ps_marmse_uneq_05_layout',
                  data = uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = ps_uneq_10, name = 'ps_marmse_uneq_10_layout',
                  data = uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = ps_uneq_15, name = 'ps_marmse_uneq_15_layout',
                  data = uneq_15, conditional = FALSE)

nps_eq_05 <- marmse.table(eq_05, 'armse.nps')
nps_eq_10 <- marmse.table(eq_10, 'armse.nps')
nps_eq_15 <- marmse.table(eq_15, 'armse.nps')

nps_uneq_05 <- marmse.table(uneq_05, 'armse.nps')
nps_uneq_10 <- marmse.table(uneq_10, 'armse.nps')
nps_uneq_15 <- marmse.table(uneq_15, 'armse.nps')

writeExcel_layout(obj_data = nps_eq_05, name = 'nps_marmse_eq_05_layout',
                  data = eq_05, conditional = FALSE)
writeExcel_layout(obj_data = nps_eq_10, name = 'nps_marmse_eq_10_layout',
                  data = eq_10, conditional = FALSE)
writeExcel_layout(obj_data = nps_eq_15, name = 'nps_marmse_eq_15_layout',
                  data = eq_15, conditional = FALSE)
writeExcel_layout(obj_data = nps_uneq_05, name = 'nps_marmse_uneq_05_layout',
                  data = uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = nps_uneq_10, name = 'nps_marmse_uneq_10_layout',
                  data = uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = nps_uneq_15, name = 'nps_marmse_uneq_15_layout',
                  data = uneq_15, conditional = FALSE)


###################
# Now do the same for the armse of the Elliott-Haviland estimator and IPF

EH_eq_05 <- marmse.table(eq_05, 'armse.EH')
EH_eq_10 <- marmse.table(eq_10, 'armse.EH')
EH_eq_15 <- marmse.table(eq_15, 'armse.EH')

EH_uneq_05 <- marmse.table(uneq_05, 'armse.EH')
EH_uneq_10 <- marmse.table(uneq_10, 'armse.EH')
EH_uneq_15 <- marmse.table(uneq_15, 'armse.EH')

writeExcel_layout(obj_data = EH_eq_05, name = 'EH_marmse_eq_05_layout',
                  data = eq_05, conditional = FALSE)
writeExcel_layout(obj_data = EH_eq_10, name = 'EH_marmse_eq_10_layout',
                  data = eq_10, conditional = FALSE)
writeExcel_layout(obj_data = EH_eq_15, name = 'EH_marmse_eq_15_layout',
                  data = eq_15, conditional = FALSE)
writeExcel_layout(obj_data = EH_uneq_05, name = 'EH_marmse_uneq_05_layout',
                  data = uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = EH_uneq_10, name = 'EH_marmse_uneq_10_layout',
                  data = uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = EH_uneq_15, name = 'EH_marmse_uneq_15_layout',
                  data = uneq_15, conditional = FALSE)

IPF_eq_05 <- marmse.table(eq_05, 'armse.IPF')
IPF_eq_10 <- marmse.table(eq_10, 'armse.IPF')
IPF_eq_15 <- marmse.table(eq_15, 'armse.IPF')

IPF_uneq_05 <- marmse.table(uneq_05, 'armse.IPF')
IPF_uneq_10 <- marmse.table(uneq_10, 'armse.IPF')
IPF_uneq_15 <- marmse.table(uneq_15, 'armse.IPF')

writeExcel_layout(obj_data = IPF_eq_05, name = 'IPF_marmse_eq_05_layout',
                  data = eq_05, conditional = FALSE)
writeExcel_layout(obj_data = IPF_eq_10, name = 'IPF_marmse_eq_10_layout',
                  data = eq_10, conditional = FALSE)
writeExcel_layout(obj_data = IPF_eq_15, name = 'IPF_marmse_eq_15_layout',
                  data = eq_15, conditional = FALSE)
writeExcel_layout(obj_data = IPF_uneq_05, name = 'IPF_marmse_uneq_05_layout',
                  data = uneq_05, conditional = FALSE)
writeExcel_layout(obj_data = IPF_uneq_10, name = 'IPF_marmse_uneq_10_layout',
                  data = uneq_10, conditional = FALSE)
writeExcel_layout(obj_data = IPF_uneq_15, name = 'IPF_marmse_uneq_15_layout',
                  data = uneq_15, conditional = FALSE)


#######
## Make plots of mean weight versus armse.comb <=> armse.ps and armse.nps

cond_equal$best <- apply(cond_equal[ , c('armse.comb','armse.nps','armse.ps')], 1,
                         which.min)
cond_equal$best <- factor(cond_equal$best, levels = 1:3, labels = c('comb','nps','ps'))
table(cond_equal$best, useNA = 'ifany')

cond_unequal$best <- apply(cond_unequal[ , c('armse.comb','armse.nps','armse.ps')], 1,
                           which.min)
cond_unequal$best <- factor(cond_unequal$best, levels = 1:3, labels = c('comb','nps','ps'))
table(cond_unequal$best, useNA = 'ifany')

cond_equal$cat.size <- 'equal categories'
cond_unequal$cat.size <- 'unequal categories'

plotdata <- rbind(
  cond_equal[ , c('cat.size','armse.comb','meanw.ps','best')],
  cond_unequal[ , c('cat.size','armse.comb','meanw.ps','best')]
)

p <- ggplot(data = plotdata, aes(x = meanw.ps, y = armse.comb, fill = best)) +
  geom_point(shape = 21) +
  scale_fill_grey(start = 0.05, end = 0.95, drop = FALSE) +
  facet_wrap(~ cat.size, ncol = 2, scales = 'free_y') +
  lims(x = c(NA, 1), y = c(0, NA)) +
  geom_hline(yintercept = 0, colour = "gray50") +
  geom_vline(xintercept = 0.65, colour = "red", linetype = 2) +
  geom_vline(xintercept = 0.75, colour = "red", linetype = 2) +
  labs(x = 'mean weight of ps',
       y = 'marmse of comb',
       fill = 'best estimator')
ggsave('Outputs/plot_comb_marmse_vs_weight.pdf',
       device = 'pdf', paper = 'a4r')


plotdata2 <- rbind(
  cond_equal[ , c('armse.ps','armse.nps','armse.comb','armse.EH','armse.IPF',
                  'bias.nps','bias.comb','bias.EH','bias.IPF','k','meanw.ps','cat.size')],
  cond_unequal[ , c('armse.ps','armse.nps','armse.comb','armse.EH','armse.IPF',
                    'bias.nps','bias.comb','bias.EH','bias.IPF','k','meanw.ps','cat.size')]
)
plotdata2$id <- 1:nrow(plotdata2)
plotdata2a <- reshape(plotdata2[ , c('id','k','meanw.ps','cat.size',
                                     'armse.ps','armse.nps','armse.comb','armse.EH','armse.IPF')],
                      idvar = c('id','k','cat.size','meanw.ps'), direction = 'long',
                      varying = c('armse.ps','armse.nps','armse.comb','armse.EH','armse.IPF'),
                      timevar = 'estimator')
plotdata2b <- reshape(plotdata2[ , c('id','k','meanw.ps','cat.size',
                                     'bias.nps','bias.comb','bias.EH','bias.IPF')],
                      idvar = c('id','k','cat.size','meanw.ps'), direction = 'long',
                      varying = c('bias.nps','bias.comb','bias.EH','bias.IPF'),
                      timevar = 'estimator')


p <- ggplot(data = plotdata2a,
       aes(x = meanw.ps, y = armse, color = estimator, shape = estimator)) +
  geom_point(alpha = 0.4) +
  facet_wrap(~ cat.size, ncol = 2, scales = 'free_y') +
  lims(x = c(NA, 1), y = c(0, NA)) +
  geom_hline(yintercept = 0, colour = "gray50") +
  labs(x = 'mean weight of ps',
       y = 'armse',
       color = 'estimator')
ggsave('Outputs/plot_allest_marmse_vs_weight.pdf',
       device = 'pdf', paper = 'a4r')

p <- ggplot(data = plotdata2b,
            aes(x = meanw.ps, y = bias, color = estimator, shape = estimator)) +
  geom_point(alpha = 0.4) +
  facet_wrap(~ cat.size, ncol = 2, scales = 'free_y') +
  lims(x = c(NA, 1)) +
  geom_hline(yintercept = 0, colour = "gray50") +
  labs(x = 'mean weight of ps',
       y = 'bias',
       color = 'estimator')
ggsave('Outputs/plot_allest_bias_vs_weight.pdf',
       device = 'pdf', paper = 'a4r')


#######
## Compute mean difference between MARMSE of combined estimator and individual estimators

diff.marmse <- matrix(0,6,2)
diff.marmse[1,1] <- mean(marmse_eq_05) - mean(ps_eq_05)
diff.marmse[1,2] <- mean(marmse_eq_05) - mean(nps_eq_05)

diff.marmse[2,1] <- mean(marmse_uneq_05) - mean(ps_uneq_05)
diff.marmse[2,2] <- mean(marmse_uneq_05) - mean(nps_uneq_05)

diff.marmse[3,1] <- mean(marmse_eq_10) - mean(ps_eq_10)
diff.marmse[3,2] <- mean(marmse_eq_10) - mean(nps_eq_10)

diff.marmse[4,1] <- mean(marmse_uneq_10) - mean(ps_uneq_10)
diff.marmse[4,2] <- mean(marmse_uneq_10) - mean(nps_uneq_10)

diff.marmse[5,1] <- mean(marmse_eq_15) - mean(ps_eq_15)
diff.marmse[5,2] <- mean(marmse_eq_15) - mean(nps_eq_15)

diff.marmse[6,1] <- mean(marmse_uneq_15) - mean(ps_uneq_15)
diff.marmse[6,2] <- mean(marmse_uneq_15) - mean(nps_uneq_15)

colnames(diff.marmse)<-c("diff comb-ps","diff comb-nps")
rownames(diff.marmse)<-c("eq_05","uneq_05",
                         "eq_10","uneq_10",
                         "eq_15","uneq_15")

writeExcel(obj = diff.marmse, name = 'summary_diff_marmse')


diff.marmse.EHandIPF <- matrix(0,6,2)
diff.marmse.EHandIPF[1,1] <- mean(marmse_eq_05 - EH_eq_05, na.rm = TRUE)
diff.marmse.EHandIPF[1,2] <- mean(marmse_eq_05) - mean(IPF_eq_05)

diff.marmse.EHandIPF[2,1] <- mean(marmse_uneq_05 - EH_uneq_05, na.rm = TRUE)
diff.marmse.EHandIPF[2,2] <- mean(marmse_uneq_05) - mean(IPF_uneq_05)

diff.marmse.EHandIPF[3,1] <- mean(marmse_eq_10 - EH_eq_10, na.rm = TRUE)
diff.marmse.EHandIPF[3,2] <- mean(marmse_eq_10) - mean(IPF_eq_10)

diff.marmse.EHandIPF[4,1] <- mean(marmse_uneq_10 - EH_uneq_10, na.rm = TRUE)
diff.marmse.EHandIPF[4,2] <- mean(marmse_uneq_10) - mean(IPF_uneq_10)

diff.marmse.EHandIPF[5,1] <- mean(marmse_eq_15 - EH_eq_15, na.rm = TRUE)
diff.marmse.EHandIPF[5,2] <- mean(marmse_eq_15) - mean(IPF_eq_15)

diff.marmse.EHandIPF[6,1] <- mean(marmse_uneq_15 - EH_uneq_15, na.rm = TRUE)
diff.marmse.EHandIPF[6,2] <- mean(marmse_uneq_15) - mean(IPF_uneq_15)

colnames(diff.marmse.EHandIPF)<-c("diff comb-EH","diff comb-IPF")
rownames(diff.marmse.EHandIPF)<-c("eq_05","uneq_05",
                                  "eq_10","uneq_10",
                                  "eq_15","uneq_15")

writeExcel(obj = diff.marmse.EHandIPF, name = 'summary_diff_marmse_EHandIPF')
