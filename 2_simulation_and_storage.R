#After creating the functions running the scripts 1_create_functions, simulations
#can be run with the first part of this script (until line 113) and the results will be stored.

#From line 116, the results to be analyzed are computed: Mean of the ARMSE
#per domain and mean absolute bias per domain

library(doParallel)

setwd('//cbsp.nl/Productie/secundair/MPOnderzoek/Werk/Combineren/Projecten/Comb_Prob_Nonprob/code_revision2')

################ Create grids to run the functions ################

conditions <- expand.grid(k=c(2,4,10,15),
                          c=c(3,5,8,15),
                          n.pk=c(20,100,400,900),
                          n.npk=c(100,1000,2000,6000),
                          age.step=c(0.5, 1, 1.5),
                          stringsAsFactors = FALSE)

###############################################
################ RUN SIMULATIONS ##############
###############################################

# cores for parallel processing
number_cores <- max(1, detectCores() - 1)
cl <- makeCluster(number_cores)
registerDoParallel(cl = cl)
if (getDoParWorkers() == 1) warning("executing sequentially, not in parallel")


####### Run simulation and storage = CONDITIONS WITH CATEGORIES OF EQUAL SIZE #######

res <- foreach(s = 1:nrow(conditions),
               .combine = c,
               .packages = c('dplyr', 'sampling', 'MASS', 'mipfp')) %dopar% {
                 
                 x <- run.estimation(k = conditions$k[s],
                                     c = conditions$c[s],
                                     n.pk = conditions$n.pk[s],
                                     n.npk = conditions$n.npk[s],
                                     equal.size = TRUE,
                                     selec.type = 'y',
                                     age.step = conditions$age.step[s],
                                     frac.domain = 0.10,
                                     model = 'age')
                 saveRDS(x, file = paste0("res_eq_", s, ".rds"))
                 
                 return(s)
               }

####### Run simulation and storage = CONDITIONS WITH CATEGORIES OF UNEQUAL SIZE #######

res <- foreach(s = 1:nrow(conditions),
               .combine = c,
               .packages = c('dplyr', 'sampling', 'MASS', 'mipfp')) %dopar% {
                 
                 x <- run.estimation(k = conditions$k[s],
                                     c = conditions$c[s],
                                     n.pk = conditions$n.pk[s],
                                     n.npk = conditions$n.npk[s],
                                     equal.size = FALSE,
                                     selec.type = 'y',
                                     age.step = conditions$age.step[s],
                                     frac.domain = 0.10,
                                     model = 'age')
                 saveRDS(x, file = paste0("res_uneq_", s, ".rds"))
                 
                 return(s)
               }

stopCluster(cl)


##################################################################################
####### Read ARMSE of each simulation and compute the mean per domain ############
####### Also compute absolute mean bias per domain                    ############
####### Also compare expected and estimated MSE of probability sample ############
##################################################################################

processResults <- function(X, cond, s) {
  
  cond$armse.ps[s] <- mean(X$armse.ps, na.rm = TRUE)
  cond$armse.nps[s] <- mean(X$armse.nps, na.rm = TRUE)
  cond$armse.comb[s] <- mean(X$armse.comb, na.rm = TRUE)
  cond$armse.EH[s] <- mean(X$armse.EH, na.rm = TRUE)
  cond$armse.IPF[s] <- mean(X$armse.IPF, na.rm = TRUE)
  
  cond$bias.nps[s] <- mean(abs(X$bias.nps), na.rm = TRUE)
  cond$bias.comb[s] <- mean(abs(X$bias.comb), na.rm = TRUE)
  cond$bias.EH[s] <- mean(abs(X$bias.EH), na.rm = TRUE)
  cond$bias.IPF[s] <- mean(abs(X$bias.IPF), na.rm = TRUE)
  
  NMADEMSE <- cond$n.pk[s] * apply(abs(X$emse.ps - X$estmse.ps), 1, mean, na.rm = TRUE)
  cond$nmademse.ps[s] <- mean(NMADEMSE, na.rm = TRUE)
  
  cond$meanratiomse.ps[s] <- mean(apply(X$emse.ps, 2, mean) / X$mse.ps[1,], na.rm = TRUE)
  cond$medianratiomse.ps[s] <- median(apply(X$emse.ps, 2, mean) / X$mse.ps[1,], na.rm = TRUE)
  
  cond$meanratiomse.nps[s] <- mean(apply(X$emse.nps, 2, mean) / X$mse.nps[1,], na.rm = TRUE)
  cond$medianratiomse.nps1[s] <- median(apply(X$emse.nps, 2, mean) / X$mse.nps[1,], na.rm = TRUE)
  
  W.ps <- X$emse.nps / (X$emse.ps + X$emse.nps)
  W.ps[W.ps < 0] <- 0
  W.ps[W.ps > 1] <- 1
  cond$meanw.ps[s] <- mean(apply(W.ps, 2, mean), na.rm = TRUE)
  cond$medianw.ps[s] <- median(apply(W.ps, 2, mean), na.rm = TRUE)
  
  return(cond)
  
}


# For equal size categories
cond.eq <- conditions

for (s in 1:nrow(cond.eq)){
  X <- readRDS(file = paste0("res_eq_", s, ".rds"))
  cond.eq <- processResults(X = X, cond = cond.eq, s = s)
}

#For unequal size categories
cond.uneq <- conditions

for (s in 1:nrow(cond.uneq)){
  X <- readRDS(file = paste0("res_uneq_", s, ".rds"))
  cond.uneq <- processResults(X = X, cond = cond.uneq, s = s)
}


# save results
saveRDS(cond.eq, file="cond_eq.rds")
saveRDS(cond.uneq, file="cond_uneq.rds")
