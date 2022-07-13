library(openxlsx)
library(dplyr)
library(data.table)
library(MASS)
library(tibble)

setwd("D:/vaccination")

df <- read.xlsx("dataset.xlsx", sheet="dataset")

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
et1.2 <- select(et1.2, -c("id", "g"))

##############Table 1
df2 <- read.xlsx("dataset.xlsx", sheet="infected cases")
df2 %>% dplyr::group_by(Age_gp, Vaccine_type, Dose) %>% 
  summarize(tot.inf = n(),
            mday.diff = mean(day.diff, na.rm=T),
            sd.day.diff = sd(day.diff, na.rm=T)) -> t1

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
ci <- as.data.frame(confint(m1, level=0.95))
ci <- tibble::rownames_to_column(ci, "coeff")
t.g1 <- merge(t.g1, ci, by = "coeff")
t.g1$rr <- exp(t.g1$Estimate)
t.g1$rr2.5 <- exp(t.g1$`2.5 %`)
t.g1$rr97.5 <- exp(t.g1$`97.5 %`)
t.g1$ve <- (1- t.g1$rr)*100
t.g1$ve2.5 <- (1- t.g1$rr2.5)*100
t.g1$ve97.5 <- (1- t.g1$rr97.5)*100

#Age group: 12-18
df.g2 <- subset(df, Age_gp == 2 & log.pop != -Inf )
m2 <- glm.nb(infection.no ~ vcdose + day.count +offset(log.pop), link = log, data=df.g2 )
t.g2 <- as.data.frame(summary(m2)$coefficients)
t.g2 <- tibble::rownames_to_column(t.g2, "coeff")
ci2 <- as.data.frame(confint(m2, level=0.95))
ci2 <- tibble::rownames_to_column(ci2, "coeff")
t.g2 <- merge(t.g2, ci2, by = "coeff")
t.g2$rr <- exp(t.g2$Estimate)
t.g2$rr2.5 <- exp(t.g2$`2.5 %`)
t.g2$rr97.5 <- exp(t.g2$`97.5 %`)
t.g2$ve <- (1- t.g2$rr)*100
t.g2$ve2.5 <- (1- t.g2$rr2.5)*100
t.g2$ve97.5 <- (1- t.g2$rr97.5)*100