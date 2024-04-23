
library(ggplot2)

setwd('//cbsp.nl/Productie/secundair/MPOnderzoek/Werk/Combineren/Projecten/Comb_Prob_Nonprob/code_revision2')

source('1_create_functions.R')

N <- 1e7

seed <- as.integer(paste0(15,3,10))
set.seed(seed)

pop05 <- create.population(k = 15, c = 3, equal.size = TRUE,
                           selec.type = 'y', age.step = 0.5,
                           frac.domain = 0.10)
pop10 <- create.population(k = 15, c = 3, equal.size = TRUE,
                           selec.type = 'y', age.step = 1.0,
                           frac.domain = 0.10)
pop15 <- create.population(k = 15, c = 3, equal.size = TRUE,
                           selec.type = 'y', age.step = 1.5,
                           frac.domain = 0.10)

lim <- sapply(levels(pop05$age), function(a) quantile(probs = c(0.025, 0.975),
                                                      x = c(subset(pop05, age == a)$yc,
                                                            subset(pop10, age == a)$yc,
                                                            subset(pop15, age == a)$yc)))

temp <- colnames(lim)
lim <- as.data.frame(t(lim))
row.names(lim) <- NULL
lim$a <- sapply(temp, function(x) which(c('Y','M','O') == x))


###

xseq <- seq(-5, 5, 0.001)

X <- data.frame(y = rep(xseq, 9),
                delta = rep(c(0.5, 1.0, 1.5), each = 3*length(xseq)),
                a = rep(rep(c(1, 2, 3), each = length(xseq)), times = 3))

bepaaldLambda <- function(input) {
  y <- input[1]
  delta <- input[2]
  a <- input[3]
  (1 / (1 + exp(-(2-delta*(a-1))*(y - 0.75))))
}

X$lambda <- apply(X, 1, bepaaldLambda)

X$delta <- factor(X$delta, levels = c(0.5, 1.0, 1.5), labels = c('0.5', '1.0', '1.5'))
X$a <- factor(X$a)
lim$a <- factor(lim$a)

###

labfun <- function(x) label_both(x, sep = " = ")

p <- ggplot(data = X,
       aes(x = y, y = lambda, linetype = delta)) +
  facet_wrap(~ a, ncol = 3, labeller = labfun) +
  geom_line() +
  geom_vline(aes(xintercept = `2.5%`), lim, colour = "red", linetype = 2) +
  geom_vline(aes(xintercept = `97.5%`), lim, colour = "red", linetype = 2) +
  geom_hline(yintercept = 0, colour = "gray50")
ggsave('Outputs/plot_lambda.pdf',
       device = 'pdf', paper = 'a4r')


###
# restcode

aux <- qnorm(p = c(0.35, 0.75), mean = 0, sd = 1, lower.tail = TRUE)
dnorm(x = c(-Inf, aux, Inf), mean = 0, sd = 1)
pnorm(q = c(-Inf, aux, Inf), mean = 0, sd = 1, lower.tail = TRUE)

E1 <- -dnorm(x = aux[1], mean = 0, sd = 1)/0.35
V1 <- 1 + aux[1]*E1 - E1^2

E2 <- - (dnorm(x = aux[2], mean = 0, sd = 1) - dnorm(x = aux[1], mean = 0, sd = 1))/0.40
V2 <- 1 - (aux[2]*dnorm(x = aux[2], mean = 0, sd = 1) - aux[1]*dnorm(x = aux[1], mean = 0, sd = 1))/0.40 - E2^2

E3 <- dnorm(x = aux[2], mean = 0, sd = 1)/0.25
V3 <- 1 + aux[2]*E3 - E3^2
