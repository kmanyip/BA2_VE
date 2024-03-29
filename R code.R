library(openxlsx)
library(dplyr)
library(data.table)
library(MASS)
library(tibble)
library(tidyverse)
library(metafor)
library(boot)

setwd("D:/vaccination")

df <- read.xlsx("submitted/dataset_v2.xlsx")

#Sex: 1= M, 2 = F
#Age group: 1: 3-11, 2= 12-18

############Extended Data table 1

#calculate total vaccination no
df %>% dplyr::group_by(Sex, Age_gp, Vaccine_type, Dose) %>% 
        summarize(vac.no = max(cum.vac.no),
                    pop = max(total.pop),
                    infection = sum(infection.no, na.rm=T)) -> et1

#calculate total vaccination no of the latest does
et1$id <- paste0(et1$Sex, et1$Age_gp, et1$Vaccine_type)
et1[order(et1$id),] -> et1
et1$g <- rleid(et1$id)

table(et1$id)

temp <- NULL
for( i in 1:max(et1$g)){
  tp <- subset(et1, g == i)
  if(NROW(tp)==3){
  tp[1,5] <- tp[1,5]-tp[2,5]
  tp[2,5] <- tp[2,5]-tp[3,5]
  }
  tp -> temp[[i]]
}

et1.1 <- do.call("rbind", temp)

et1.1$id <- paste0(et1.1$Sex, et1.1$Age_gp)
et1.1[order(et1.1$id),] -> et1.1
et1.1$g <- rleid(et1.1$id)

table(et1.1$id)

temp <- NULL
for( i in 1:max(et1.1$g)){
  tp <- subset(et1.1, g == i)
  tp$tot.inf <- sum(tp$infection, na.rm=T)
  tp$vac.per <- tp$vac.no/tp$pop*100
  tp$inf.per <- tp$infection/tp$tot.inf*100
  tp -> temp[[i]]
}

et1.2 <- do.call("rbind", temp)
et1.2 <- dplyr::select(et1.2, -c("id", "g"))

##############Table 1
####create variables to calculate overall mean and sd of days elapsed from last dose
df$day.diff <- df$mean.day.diff*df$infection.no
df$nsd <- (df$infection.no-1)*df$sd.day.diff^2

df %>% dplyr::group_by(Age_gp, Vaccine_type, Dose) %>% 
  summarize(tot.inf = sum(infection.no, na.rm=T),
            mday.diff = sum(day.diff, na.rm=T)/sum(infection.no, na.rm=T),
            sd.day.diff = sqrt(sum(nsd, na.rm=T)/(sum(infection.no, na.rm=T)-14))) -> t1

############Vaccine effectiveness
df$log.pop <- log(df$pop.at.risk)
df$vcdose <- with(df, ifelse(Vaccine_type == "BioNTech" & Dose ==1, 1, ifelse(
  Vaccine_type == "Sinovac" & Dose ==1, 2, ifelse(
    Vaccine_type == "BioNTech" & Dose ==2, 3, ifelse( 
      Vaccine_type == "Sinovac" & Dose ==2, 4, ifelse(
        Vaccine_type == "BioNTech" & Dose ==3, 5, ifelse(
          Vaccine_type == "Sinovac" & Dose ==3, 6, 0)))))))

df[order(df$Date),] -> df

df$day.count <- rleid(df$Date)

df$vcdose <- as.factor(df$vcdose)

#Age group: 3-11
df.g1 <- subset(df, Age_gp == 1 & log.pop != -Inf & (vcdose ==0|vcdose == 1 | vcdose ==2 | vcdose ==4))
m1 <- glm.nb(infection.no ~ vcdose + day.count +offset(log.pop), link = log, data=df.g1 )
t.g1 <- as.data.frame(summary(m1)$coefficients)
t.g1 <- tibble::rownames_to_column(t.g1, "coeff")
t.g1$rr <- exp(t.g1$Estimate)
t.g1$ve <- (1- t.g1$rr)*100
ci <- as.data.frame(confint(m1, level=0.95))
ci <- tibble::rownames_to_column(ci, "coeff")
t.g1 <- merge(t.g1, ci, by = "coeff")
t.g1$rr <- exp(t.g1$Estimate)
t.g1$rr.se <- exp(t.g1$`Std. Error`)
t.g1$rr2.5 <- exp(t.g1$`2.5 %`)
t.g1$rr97.5 <- exp(t.g1$`97.5 %`)
t.g1$ve <- (1- t.g1$rr)*100
t.g1$ve2.5 <- (1- t.g1$rr2.5)*100
t.g1$ve97.5 <- (1- t.g1$rr97.5)*100

t.g1.2 <- dplyr::select(t.g1, c("coeff", "rr"))
t.g1.2$Age_gp = "1"

#Age group: 12-18
df.g2 <- subset(df, Age_gp == 2 & log.pop != -Inf )
m2 <- glm.nb(infection.no ~ vcdose + day.count +offset(log.pop), link = log, data=df.g2)
t.g2 <- as.data.frame(summary(m2)$coefficients)
t.g2 <- tibble::rownames_to_column(t.g2, "coeff")
t.g2$rr <- exp(t.g2$Estimate)
t.g2$ve <- (1- t.g2$rr)*100
ci2 <- as.data.frame(confint(m2, level=0.95))
ci2 <- tibble::rownames_to_column(ci2, "coeff")
t.g2 <- merge(t.g2, ci2, by = "coeff")
t.g2$rr <- exp(t.g2$Estimate)
t.g2$rr.se <- exp(t.g2$`Std. Error`)
t.g2$rr2.5 <- exp(t.g2$`2.5 %`)
t.g2$rr97.5 <- exp(t.g2$`97.5 %`)
t.g2$ve <- (1- t.g2$rr)*100
t.g2$ve2.5 <- (1- t.g2$rr2.5)*100
t.g2$ve97.5 <- (1- t.g2$rr97.5)*100

t.g2.2 <- dplyr::select(t.g2, c("coeff", "rr"))
t.g2.2$Age_gp = "2"

rr <- rbind(t.g1.2, t.g2.2)
colnames(rr) <- c("vcdose", "rr", "Age_gp")

## Bootstrapped CI
#Age group: 3-11
fun <- function(df.g1, i){
m1 <- glm.nb(infection.no ~ vcdose + day.count +offset(log.pop), link = log, data=df.g1[i,])
bt.g1 <- as.data.frame(summary(m1)$coefficients)
bt.g1 <- tibble::rownames_to_column(bt.g1, "coeff")
bt.g1$rr <- exp(bt.g1$Estimate)
bt.g1$ve <- (1- bt.g1$rr)*100
ci <- as.data.frame(confint(m1, level=0.95))
ci <- tibble::rownames_to_column(ci, "coeff")
bt.g1 <- merge(bt.g1, ci, by = "coeff")
bt.g1$rr <- exp(bt.g1$Estimate)
bt.g1$rr.se <- exp(bt.g1$`Std. Error`)

bt.g1.2 <- dplyr::select(bt.g1, c("coeff", "rr", "rr.se"))
bt.g1.2$Age_gp = "1"

bt.g1.3 <- bt.g1.2[3:5,]
bt.g1.3 <- subset(bt.g1.3, coeff != "vcdose2")
bt.g1.3$rr <- log(bt.g1.3$rr)
bt.g1.3$rr.se <- log(bt.g1.3$rr.se)
rs<- rma(rr, rr.se, method='FE', measure='PR',data=bt.g1.3)
g1.rr.pooled <- NULL
g1.rr.pooled <- predict(rs, transf=exp, digits=3)$pred

return(pooled.rr1 = g1.rr.pooled)
}

rr.g1 <- boot(df.g1, fun, R=1000)

rr.g1.2 <- as.data.frame(t(boot.ci(rr.g1, type="norm")$normal[,2:3]))
rr.g1.2$vcdose <- "pooled.rr"

colnames(rr.g1.2) <- c("ci.lb", "ci.ub", "vcdose")

rr.g1.2$Age_gp <- 1

#Age group: 12-18
fun <- function(df.g2, i){
m2 <- glm.nb(infection.no ~ vcdose + day.count +offset(log.pop), link = log, data=df.g2[i,])
bt.g2 <- as.data.frame(summary(m2)$coefficients)
bt.g2 <- tibble::rownames_to_column(bt.g2, "coeff")
ci2 <- as.data.frame(confint(m2, level=0.95))
ci2 <- tibble::rownames_to_column(ci2, "coeff")
bt.g2 <- merge(bt.g2, ci2, by = "coeff")
bt.g2$rr <- exp(bt.g2$Estimate)
bt.g2$rr.se <- exp(bt.g2$`Std. Error`)

bt.g2.2 <- dplyr::select(bt.g2, c("coeff", "rr", "rr.se"))
bt.g2.2$Age_gp = "2"

bt.g2.3 <- bt.g2.2[3:8,]
###use VE of dose 2 to replace dose 3
bt.g2.3$rr[5] <- bt.g2.3$rr[3]
bt.g2.3$rr[6] <- bt.g2.3$rr[4]
bt.g2.3$rr.se[5] <- bt.g2.3$rr.se[3]
bt.g2.3$rr.se[6] <- bt.g2.3$rr.se[4]
bt.g2.3$rr <- log(bt.g2.3$rr)
bt.g2.3$rr.se <- log(bt.g2.3$rr.se)

rs<- rma(rr, rr.se, method='FE', measure='PR',data=bt.g2.3)
g2.rr.pooled <- NULL
g2.rr.pooled <- predict(rs, transf=exp, digits=3)$pred

return(pooled.rr1 = g2.rr.pooled)
}

rr.g2 <- boot(df.g2, fun, R=1000)

rr.g2.2 <- as.data.frame(t(boot.ci(rr.g2, type="norm", index=1)$normal[,2:3]))
rr.g2.2$vcdose <- "pooled.rr"

colnames(rr.g2.2) <- c("ci.lb", "ci.ub", "vcdose")

rr.g2.2$Age_gp <- 2

pooled.rr <- dplyr::select(rbind(rr.g1.2, rr.g2.2), - "vcdose")

# pooled.rr <- dplyr::select(subset(rr, vcdose == "pooled.rr"), c("Age_gp", "ci.lb", "ci.ub"))
# vcdose <- dplyr::select(subset(rr, vcdose != "pooled.rr"), c("vcdose", "rr", "Age_gp"))

#####################daily expected number
df %>% dplyr::group_by(Date, Age_gp,vcdose) %>% summarize(tot.inf = sum(infection.no, na.rm=T)) -> ep.case
ep.case$vcdose <- paste0("vcdose", ep.case$vcdose)

ep.case2 <- merge(ep.case, rr, by = c("vcdose", "Age_gp"), all.x=T)
ep.case2$rr <- with(ep.case2, ifelse(is.na(rr),1, rr))

ep.case2 <- merge(ep.case2, pooled.rr, by = "Age_gp", all.x=T)

ep.case2 %>% gather(var, value, tot.inf:rr) -> ep.case4
ep.case4$id <- paste0(ep.case4$vcdose, ep.case4$var)

ep.case4 %>% dplyr::select(-c("var", "vcdose")) %>% tidyr::spread(id, value) -> ep.case5

ep.case5$tot.case <- with(ep.case5, ifelse(Age_gp ==1, vcdose0tot.inf+vcdose1tot.inf + vcdose2tot.inf + vcdose4tot.inf, ifelse(Age_gp ==2, 
                                                   vcdose0tot.inf+vcdose1tot.inf + vcdose2tot.inf+ vcdose3tot.inf
                                                     +vcdose4tot.inf+ vcdose5tot.inf + vcdose6tot.inf,NA )))

ep.case5$ep.case <- with(ep.case5, ifelse(Age_gp ==1, vcdose0tot.inf+ vcdose2tot.inf+ vcdose1tot.inf/vcdose1rr + vcdose4tot.inf/vcdose4rr, ifelse(Age_gp ==2, 
                                  vcdose0tot.inf+ vcdose1tot.inf/vcdose1rr  + vcdose2tot.inf/vcdose2rr + (vcdose3tot.inf+ vcdose5tot.inf )/vcdose3rr
                                  + (vcdose4tot.inf+ vcdose6tot.inf)/vcdose4rr ,NA )))


ep.case5$ep.case97.5 <- with(ep.case5, ifelse(Age_gp ==1, vcdose0tot.inf+ vcdose2tot.inf + (vcdose1tot.inf +  vcdose4tot.inf)/ci.lb, ifelse(Age_gp ==2, 
                           vcdose0tot.inf+(vcdose1tot.inf + vcdose2tot.inf + vcdose3tot.inf
                        +vcdose4tot.inf+ vcdose5tot.inf + vcdose6tot.inf)/ci.lb,NA )))


ep.case5$ep.case2.5 <- with(ep.case5, ifelse(Age_gp ==1, vcdose0tot.inf+ vcdose2tot.inf + (vcdose1tot.inf + vcdose4tot.inf)/ci.ub, ifelse(Age_gp ==2, 
                         vcdose0tot.inf+(vcdose1tot.inf + vcdose2tot.inf + vcdose3tot.inf
                                  +vcdose4tot.inf+ vcdose5tot.inf + vcdose6tot.inf)/ci.ub,NA )))


                                      
