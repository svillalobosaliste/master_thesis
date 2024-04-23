#This script create the functions to be able to run the simulations

##########Loading packages #############
library(dplyr)
library(sampling)
library(MASS)
library(mipfp)

N <- 100000 #Sample size of the population
rep <- 1000 #Number of simulations

###########################################################################
# Elliott and Haviland (EH) estimator
###########################################################################
EH_estimator <- function(f_Z_P, f_Z_NP, f_n_P, f_n_NP) {
  f_sigma2_P <- f_Z_P*(1 - f_Z_P)
  f_sigma2_NP <- f_Z_NP*(1 - f_Z_NP)
  f_bias <- f_Z_NP - f_Z_P
  f_estimate <- (f_Z_NP*(f_sigma2_P/f_n_P) + f_Z_P*(f_bias^2 + f_sigma2_NP/f_n_NP))/(f_bias^2 + (f_sigma2_NP/f_n_NP) + (f_sigma2_P/f_n_P))
  return(f_estimate)
}

###########################################################################
# IPF estimator
###########################################################################
IPF_estimator <- function(i_table_NP, i_row_totals_P, i_col_totals_P) {
  i_target_data <- list(i_col_totals_P, i_row_totals_P)
  i_target_list <- list(2,1) # order of totals
  i_estimate <- Ipfp(i_table_NP, i_target_list, i_target_data)
  return(i_estimate)
}

###########################################################################

############### SAMPLING  FUNCTION ###############
#This function generates a probability and a non-probability sample.
#It is used inside the other functions.

generate.sample <- function(
    #Randomly generates PS with equal inclusion
  #Systemic random samples NPS
  population, size_ps = size_ps, size_nps = size_nps,k=k){
  PS<-NULL
  NPS<-NULL #starts with empty 
  for (i in (1:k)) {
    domain <- population[population$dom==i,] #select cases that belong to i domain
    PS <- rbind(PS,domain[sample.int(nrow(domain), size_ps),]) #add domains
    lam <- domain$lambda
    lam <- inclusionprobabilities(lam,size_nps)
    NPS <- rbind(NPS,domain[as.logical(UPrandomsystematic(lam)), ])
  }
  PS$sample <- 1
  NPS$sample <- 0
  combined_sample <- full_join(PS, NPS)
  return(combined_sample)
}

create.population <- function(
    k,              # number of domains
    c,              # number of categories
    equal.size,     # categories of equal size (TRUE) or unequal size (FALSE)
    selec.type,     # type of selectivity mechanism:
                    #   selec.type = 'y' for dependence on target variable y
    age.step = 0,   # step size to define effect age on selection probabilities lambda (default = 0)
    frac.domain = 0 # fraction of variance in age (continuous) contributed by domain-specific effect (default = 0)
){
  
  pop <- mvrnorm(N,c(0,0),diag(2)) #create population of 2 variables
  d <- sqrt(1/3)
  pop <- as.data.frame(pop)
  names(pop) <- c("X1","X2")
  pop$dom <- as.factor(sample(1:k,N,replace=T))
  
  if (frac.domain > 0) {
    # add domain-specific effect on age
    for (x in 1:k){
      pop$X1[pop$dom==x] <- (1 - sqrt(frac.domain)) * pop$X1[pop$dom==x] + rnorm(1,0,sqrt(frac.domain)) # every unit inside domain x gets the same contribution
    }
  }
  
  pop$age <- cut(pop$X1, breaks = quantile(pop$X1, probs = c(0, 0.35, 0.75, 1)),
                 labels = c('Y','M','O'), include.lowest = TRUE)
  
  pop$yc <- d*pop$X1 + d*pop$X2 + rnorm(N,0,d)
  
  if (equal.size) {
    br <- quantile(pop$yc,probs = (0:(c))/c)
  } else {
    br<-NA     
    if (c=="3"){
      br=c(0,.29,.65,1)
    } else if (c=="5") {
      br=c(0,.09,.29,.65,.87,1) 
    } else if (c=="8") {
      br=c(0,.09,.20,.28,.42,.55,.65,.87,1) 
    } else if (c=="15") {
      br=c(0,.03,.06,.11,.18,.27,.37,.49,.58,.68,.77,.85,.93,.95,.97,1) 
    }
    br <- quantile(pop$yc,probs=br)
  }
  
  pop$y <- cut(pop$yc,breaks=br,labels=1:c,include.lowest = TRUE)
  
  if (selec.type == 'y') {
    for (a in levels(pop$age)){
      pop$lambda[pop$age == a] <- 1/(1+exp(-((2 - age.step*(a == 'M') - 2*age.step*(a == 'O'))*(pop$yc[pop$age == a] - 0.75))))
    }
  } else {
    stop('Unknown type of selectivity mechanism!')
  }
  
  
  return(pop)
}

##################################################

#Elements that will be saved in a list by the next function:

#1 = estimates
#2 = mean squared error of probability sample
#3 = mean squared error of non-probability sample
#4 = mean squared error of combined estimator
#5 = mean squared error of Elliott and Haviland's estimator
#6 = mean squared error of IPF estimator
#7 = average root mean squared error of probability sample 
#8 = average root mean squared error of non-probability sample 
#9 = average root mean squared error of combined estimator
#10 = average root mean squared error of Elliott and Haviland's estimator
#11 = average root mean squared error of IPF estimator
#12 = bias of non-probability sample 
#13 = bias of combined estimator
#14 = bias of Elliott and Haviland's estimator
#15 = bias of IPF estimator
#16 = expected mean squared error of probability sample
#17 = expected mean squared error of non-probability sample
#18 = direct estimate of mean squared error of probability sample

run.estimation <- function(
    k,                # number of domains
    c,                # number of categories
    n.pk,             # size of probability sample per domain
    n.npk,            # size of non-probability sample per domain
    equal.size,       # categories of equal size (TRUE) or unequal size (FALSE)
    selec.type,       # type of selectivity mechanism:
                      #   selec.type = 'y' for dependence on target variable y
    age.step = 0,     # step size to define effect age on selection probabilities lambda (default = 0)
    frac.domain = 0,  # fraction of variance in age (continuous) contributed by domain-specific effect (default = 0)
    model = NULL      # predictors to be used in model for bias in combined estimator
){
  seed <- as.integer(paste0(k,c,10))
  set.seed(seed)
  
  pop <- create.population(k = k, c = c, equal.size = equal.size,
                           selec.type = selec.type, age.step = age.step,
                           frac.domain = frac.domain)
  
  estim <- matrix(0,rep,k*c*6)
  
  mse.ps <- matrix(0,1,k*c)
  mse.nps <- matrix(0,1,k*c)
  mse.comb <- matrix(0,1,k*c)
  mse.EH <- matrix(0,1,k*c)
  mse.IPF <- matrix(0,1,k*c)
  
  bias.nps <- matrix(0,1,k*c)
  bias.comb <- matrix(0,1,k*c)
  bias.EH <- matrix(0,1,k*c)
  bias.IPF <- matrix(0,1,k*c)
  
  # extra output
  emse.ps <- matrix(0,rep,k*c)
  emse.nps <- matrix(0,rep,k*c)
  estmse.ps <- matrix(0,rep,k*c)
  
  for(i in 1:rep){
    samp <- generate.sample(pop,n.pk,n.npk,k)
    prob <- samp[which(samp$sample=="1"),]
    np <- samp[which(samp$sample=="0"),]
    probtab <- prop.table(table(prob$dom,prob$y),margin=1)
    nprobtab <- prop.table(table(np$dom,np$y),margin=1)
    true <- prop.table(table(pop$dom,pop$y),margin=1)
    
    # compute combined estimator
    v.kc <- nprobtab*(1-nprobtab)
    
    if (is.null(model)) {
      bc <- matrix(colMeans(nprobtab-probtab),nrow=k,ncol=c,byrow=T)
    } else {
      probtab.strat <- prop.table(table(prob[ , c('dom','y',model)]),margin=c(1,3))
      nprobtab.strat <- replicate(n = dim(probtab.strat)[3], nprobtab)
      
      nk.strat <- prop.table(table(pop[(pop[ , model] %in% dimnames(probtab.strat)[[3]]), c('dom',model)]), margin = 1)
      bc.strat <- apply(nprobtab.strat - probtab.strat, 2, weighted.mean, w = nk.strat, na.rm = TRUE)
      bc.strat[is.na(bc.strat)] <- 0
      bc <- matrix(bc.strat,nrow=k,ncol=c,byrow=T)
    }
    
    if (k > 1) {
      sigma2 <- sum((nprobtab - probtab - bc)^2) / ((k-1)*c)
    } else {
      sigma2.k <- (n.pk/(n.pk-1))*((nprobtab-probtab)^2-(1-1/n.pk)*(bc^2)-(2/n.pk)*bc*nprobtab)-v.kc*(1+n.npk/n.pk)/(n.npk-1)
      sigma2.k <- apply(sigma2.k,MARGIN = 1,FUN = mean)
      sigma2 <- matrix(sigma2.k,nrow=k,ncol=c) 
    }
    
    wkc <- ((n.npk-1)*(bc^2+sigma2)+v.kc)/((n.npk-1)*(1-1/n.pk)*(bc^2+sigma2)+(1+n.npk/n.pk)*v.kc+((n.npk-1)/n.pk)*bc*(2*nprobtab-1))
    wkc[wkc<0] <- 0
    wkc[wkc>1] <- 1
    ccc <- wkc*probtab + (1-wkc)*nprobtab
    # rescale
    ccc <- ccc / matrix(rowSums(ccc), nrow = k, ncol = c, byrow = FALSE)
    
    # compute estimator by Elliott and Haviland
    EH <- EH_estimator(f_Z_P = probtab, f_Z_NP = nprobtab, f_n_P = n.pk, f_n_NP = n.npk)
    # rescale
    EH <- EH / matrix(rowSums(EH), nrow = k, ncol = c, byrow = FALSE)
    
    # compute IPF estimator
    IPF <- IPF_estimator(i_table_NP = nprobtab, i_row_totals_P = rowSums(probtab), i_col_totals_P = colSums(probtab))
    
    estim[i,1:(k*c)] <- c(t(probtab))            #PROBABILITY SAMPLE
    estim[i,(k*c+1):(2*k*c)] <- c(t(nprobtab))   #NON-PROBABILITY SAMPLE
    estim[i,(2*k*c+1):(3*k*c)] <- c(t(ccc))      #COMBINED ESTIMATOR
    estim[i,(3*k*c+1):(4*k*c)] <- c(t(EH))       #ELLIOTT & HAVILAND ESTIMATOR
    estim[i,(4*k*c+1):(5*k*c)] <- c(t(IPF$x.hat))#IPF ESTIMATOR
    estim[i,(5*k*c+1):(6*k*c)] <- c(t(true))     #TRUE VALUES
    
    mse.ps[1,1:(k*c)] <- mse.ps[1,1:(k*c)] + ((estim[i,1:(k*c)]-estim[i,(5*k*c+1):(6*k*c)])^2)/rep
    mse.nps[1,1:(k*c)] <- mse.nps[1,1:(k*c)] + ((estim[i,(k*c+1):(2*k*c)]-estim[i,(5*k*c+1):(6*k*c)])^2)/rep
    mse.comb[1,1:(k*c)] <- mse.comb[1,1:(k*c)] + ((estim[i,(2*k*c+1):(3*k*c)]-estim[i,(5*k*c+1):(6*k*c)])^2)/rep
    mse.EH[1,1:(k*c)] <- mse.EH[1,1:(k*c)] + ((estim[i,(3*k*c+1):(4*k*c)]-estim[i,(5*k*c+1):(6*k*c)])^2)/rep
    mse.IPF[1,1:(k*c)] <- mse.IPF[1,1:(k*c)] + ((estim[i,(4*k*c+1):(5*k*c)]-estim[i,(5*k*c+1):(6*k*c)])^2)/rep
    
    bias.nps[1,1:(k*c)] <- bias.nps[1,1:(k*c)] + (estim[i,(k*c+1):(2*k*c)]-estim[i,(5*k*c+1):(6*k*c)])/rep
    bias.comb[1,1:(k*c)] <- bias.comb[1,1:(k*c)] + (estim[i,(2*k*c+1):(3*k*c)]-estim[i,(5*k*c+1):(6*k*c)])/rep
    bias.EH[1,1:(k*c)] <- bias.EH[1,1:(k*c)] + (estim[i,(3*k*c+1):(4*k*c)]-estim[i,(5*k*c+1):(6*k*c)])/rep
    bias.IPF[1,1:(k*c)] <- bias.IPF[1,1:(k*c)] + (estim[i,(4*k*c+1):(5*k*c)]-estim[i,(5*k*c+1):(6*k*c)])/rep
    
    # extra output
    emse.ps[i, ] <- (1/n.pk) * ((n.npk/(n.npk-1))*v.kc + bc*(2*nprobtab-1) - bc^2 - sigma2)
    emse.nps[i, ] <- bc^2 + sigma2 + v.kc / (n.npk-1)
    estmse.ps[i, ] <- probtab*(1-probtab)/n.pk
  }
  
  rmse.ps <- sqrt(mse.ps)
  rmse.nps <- sqrt(mse.nps)
  rmse.comb <- sqrt(mse.comb)
  rmse.EH <- sqrt(mse.EH)
  rmse.IPF <- sqrt(mse.IPF)
  
  armse.ps <- matrix(0,1,k)
  armse.nps <- matrix(0,1,k)
  armse.comb <- matrix(0,1,k)
  armse.EH <- matrix(0,1,k)
  armse.IPF <- matrix(0,1,k)
  
  for (i in 1:k) {
    armse.ps[1,i] <- mean(rmse.ps[1,(i*c-c+1):(c*i)])
    armse.nps[1,i] <- mean(rmse.nps[1,(i*c-c+1):(c*i)])
    armse.comb[1,i] <- mean(rmse.comb[1,(i*c-c+1):(c*i)])
    armse.EH[1,i] <- mean(rmse.EH[1,(i*c-c+1):(c*i)])
    armse.IPF[1,i] <- mean(rmse.IPF[1,(i*c-c+1):(c*i)])
  }
  
  return(list(estim = estim,
              mse.ps = mse.ps,
              mse.nps = mse.nps,
              mse.comb = mse.comb,
              mse.EH = mse.EH,
              mse.IPF = mse.IPF,
              armse.ps = armse.ps,
              armse.nps = armse.nps,
              armse.comb = armse.comb,
              armse.EH = armse.EH,
              armse.IPF = armse.IPF,
              bias.nps = bias.nps,
              bias.comb = bias.comb,
              bias.EH = bias.EH,
              bias.IPF = bias.IPF,
              emse.ps = emse.ps,
              emse.nps = emse.nps,
              estmse.ps = estmse.ps))
  
}




#####
## Writing output to Excel


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


# function to write output table as a formatted Excel file with openxlsx
writeExcel_layout <- function(obj_data, obj_color = NULL, name, data, conditional = FALSE) {
  # obj_data: name of R object to be written as output
  # obj_color: name of R object to be used for layout (only used if conditional == TRUE)
  # name: intended name of Excel file (without .xlsx extension)
  # data: source data used to create obj_data and obj_color
  # conditional: conditional layout added?
  
  K_vals <- sort(unique(data$k))
  C_vals <- sort(unique(data$c))
  p_vals <- sort(unique(data$n.pk))
  np_vals <- sort(unique(data$n.npk))
  
  wb <- createWorkbook()
  addWorksheet(wb, name)
  writeData(wb, name, obj_data,
            xy = c(3, 3), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name,
            t(do.call(c, lapply(C_vals, function(s) c(paste0('c = ', s), rep(NA, length(np_vals) - 1))))),
            xy = c(3, 2), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name,
            do.call(c, lapply(K_vals, function(s) c(paste0('K = ', s), rep(NA, length(p_vals) - 1)))),
            xy = c(2, 3), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name, 'PS',
            xy = c((ncol(obj_data) + 3), 2), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name,
            rep(p_vals, times = length(K_vals)),
            xy = c(ncol(obj_data) + 3, 3), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name, 'NPS',
            xy = c(2, nrow(obj_data) + 3), rowNames = FALSE, colNames = FALSE)
  writeData(wb, name,
            t(rep(np_vals, times = length(C_vals))),
            xy = c(3, nrow(obj_data) + 3), rowNames = FALSE, colNames = FALSE)
  
  for (i in 1:length(C_vals)) {
    mergeCells(wb, name, cols = 2 + (i-1)*length(np_vals) + 1:length(np_vals), rows = 2)
  }
  for (i in 1:length(K_vals)) {
    mergeCells(wb, name, rows = 2 + (i-1)*length(p_vals) + 1:length(p_vals), cols = 2)
  }
  
  addStyle(wb, name, rows = 2:(nrow(obj_data) + 3), cols = 2:(ncol(obj_data) + 3),
           style = createStyle(fontSize = 8, halign = 'center', valign = 'center'),
           gridExpand = TRUE)
  addStyle(wb, name, rows = nrow(obj_data) + 3, cols = 3:(ncol(obj_data) + 2),
           style = createStyle(fontSize = 8, halign = 'center', valign = 'center', textRotation = 90))
  addStyle(wb, name, rows = 2, cols = 2:(ncol(obj_data) + 3),
           style = createStyle(fontSize = 8, halign = 'center', valign = 'center',
                               border = c('top', 'bottom')),
           gridExpand = TRUE)
  addStyle(wb, name, rows = 3:(nrow(obj_data) + 2), cols = 3:(ncol(obj_data) + 2),
           style = createStyle(fontSize = 8, halign = 'center', valign = 'center', numFmt = '0.00'),
           gridExpand = TRUE)
  for (i in 1:length(K_vals)) {
    addStyle(wb, name, rows = 2 + i*length(p_vals), cols = c(2, ncol(obj_data) + 3),
             style = createStyle(fontSize = 8, halign = 'center', valign = 'center',
                                 border = 'bottom'),
             gridExpand = TRUE)
    addStyle(wb, name, rows = 2 + i*length(p_vals), cols = 3:(ncol(obj_data) + 2),
             style = createStyle(fontSize = 8, halign = 'center', valign = 'center', numFmt = '0.00',
                                 border = 'bottom'),
             gridExpand = TRUE)
  }
  
  setColWidths(wb, name, 2, widths = 4.71)
  setColWidths(wb, name, 3:(ncol(obj_data)+3), widths = 3.57)
  
  if (conditional) {
    # conditional formatting based on colors
    writeData(wb, name, obj_color,
              xy = c(21, 3), rowNames = FALSE, colNames = FALSE)
    conditionalFormatting(wb, name, rows = 3:(nrow(obj_data) + 2), cols = 3:(ncol(obj_data) + 2),
                          rule = 'U3 == 2',
                          style = createStyle(fontSize = 8, halign = 'center', valign = 'center', numFmt = '0.00',
                                              bgFill = '#DBDBDB'))
    conditionalFormatting(wb, name, rows = 3:(nrow(obj_data) + 2), cols = 3:(ncol(obj_data) + 2),
                          rule = 'U3 == 3',
                          style = createStyle(fontSize = 8, halign = 'center', valign = 'center', numFmt = '0.00',
                                              bgFill = '#AEAAAA'))
    conditionalFormatting(wb, name, rows = 3:(nrow(obj_data) + 2), cols = 3:(ncol(obj_data) + 2),
                          rule = 'U3 == 6',
                          style = createStyle(fontSize = 8, halign = 'center', valign = 'center', numFmt = '0.00',
                                              bgFill = '#757171'))
  }
  
  saveWorkbook(wb, file = paste0('Outputs/Final tables/', name, '.xlsx'),
               overwrite = TRUE)
}
