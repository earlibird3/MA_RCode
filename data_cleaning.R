# =============================================================================
# University of St.Gallen
# Course: Master Thesis
# Authors: Michele Schoch
# Professor: Dr. Simon Pfister
# Date: 22.08.2017
# =============================================================================


###############################################################################
# =============================================================================
# Set working directory & Install Packages
# =============================================================================
###############################################################################

#set working directory
#setwd('C:/Users/Michèle/Dropbox/Master/MA/06_MA_Files/MA_RCode')
#setwd('C:/Users/Helena Aebersold/Dropbox/Michele_MA/05_ClusterAnalysis')

#install packagae 'pacman' in order to install all packages which have not been installed yet
#storage C:\Users\Michèle\Appdata_raw\Local\Temp\RtmpQxfK53\downloaded_packages
if (!require("pacman")) install.packages("pacman")
#p_load: wraps library and require into one function, it checks if a package is installed, if not it installs the package
pacman::p_load("ggplot2", "xlsx", "rJava", "xlsxjars", "reshape", "vars", "MuMIn", "tidyr", "plyr", "gridExtra",
              "lubridate", "grid")



###############################################################################
# ============================================================================
# Import data_raw file
# =============================================================================
############################################################################### 
rm(list = ls())

#file including factors worked out with Buhler group
kpi <- read.csv("06_data.csv", header=T, sep=";", stringsAsFactors=F)
colnames(kpi)[1] <- "BPMID"

#Age Ten PM AM
ageten <-read.csv("07_Age_Ten_PM_AM.csv", header=T, sep=";", stringsAsFactors=F)
colnames(ageten) <- c("BPMID", "PMAge2", "PMTen2", "AMAge2", "AMTen2")
ageten[] <- lapply(ageten, function(x) as.numeric(as.character(x)))

#TO and DB1 data_raw
to <- read.csv("TO_Act.csv", header=T, sep=";", stringsAsFactors=F)
to$ProjectID <- NULL
colnames(to) <- c("BPMID", "TOAct", "DB1Budabs", "DB1Actabs")
to[] <- lapply(to, function(x) as.numeric(as.character(x)))

#Region data_raw
region <- read.csv("Region.csv", header=T, sep=";", stringsAsFactors=F)
region$Country <- NULL
colnames(region) <- c("BPMID", "Region")

#merge to and kpi and region
data_raw <- merge(kpi,region, "BPMID", stringAsFactors = F)
#merge data_raw and to
data_raw <- merge(data_raw,to, "BPMID",stringAsFactors = F)
#merge data_raw and ageten
data_raw <- merge(data_raw,ageten, "BPMID",stringAsFactors = F)

#test if Age and Ten for PM and AM are the same except for NA
testageten <- data.frame(data_raw$BPMID,
                         data_raw$PMAge, data_raw$PMAge2, testpmage <- data_raw$PMAge-data_raw$PMAge2,
                         data_raw$PMTen, data_raw$PMTen2, testpmten <- data_raw$PMTen-data_raw$PMTen2,
                         data_raw$AMAge, data_raw$AMAge2, testamage <- data_raw$AMAge-data_raw$AMAge2,
                         data_raw$AMTen, data_raw$AMTen2, testamten <- data_raw$AMTen-data_raw$AMTen2)
max(testageten$testpmage....data_raw.PMAge...data_raw.PMAge2, na.rm=T)
max(testageten$testpmten....data_raw.PMTen...data_raw.PMTen2, na.rm=T)
max(testageten$testamage....data_raw.AMAge...data_raw.AMAge2, na.rm=T)
max(testageten$testamten....data_raw.AMTen...data_raw.AMTen2, na.rm=T)

min(testageten$testpmage....data_raw.PMAge...data_raw.PMAge2, na.rm=T)
min(testageten$testpmten....data_raw.PMTen...data_raw.PMTen2, na.rm=T)
min(testageten$testamage....data_raw.AMAge...data_raw.AMAge2, na.rm=T)
min(testageten$testamten....data_raw.AMTen...data_raw.AMTen2, na.rm=T)

write.xlsx(testageten, "testageten.xlsx")

#set old columns PMAge/Ten and AMAge/Ten NULL
data_raw$PMAge <- NULL
data_raw$PMTen <- NULL
data_raw$AMAge <- NULL
data_raw$AMTen <- NULL


#write data_raw to excle
write.xlsx(data_raw, "01_data_raw_Ro.xlsx")



###############################################################################
# ============================================================================
# Format data_raw and delete in vain columns
# =============================================================================
############################################################################### 

#read CuNo as character
data_raw$CuNo <- as.character(data_raw$CuNo)

#read ConPart as logical
data_raw$ConPart[data_raw$ConPart=='X'] <- TRUE
data_raw$ConPart[data_raw$ConPart=='#'] <- FALSE
data_raw$ConPart <- as.character(data_raw$ConPart) #logical is disturbing melt

#eliminate columns
data_raw$monthsbetw.HOMand1strcstb <- NULL #calculating vectors
data_raw$monthsbetw.HOMand1strqtyb <- NULL
data_raw$monthsbetw.HOMand1strtimeb <- NULL
data_raw$HOMvsyellow.redstatuscostsb <- NULL
data_raw$HOMvsyellow.redstatusqualityb <- NULL
data_raw$HOMvsyellow.redstatustimeb <- NULL
data_raw$Medianofavg.BAprojectb <- NULL
data_raw$Medianofavg.BUprojectb <- NULL
data_raw$Medianofavg.MSprojectb <- NULL
data_raw$ORDate <- NULL #is same as HOM, not used in analysis
#data_raw$BPMID <- NULL #only identifies project, not used in analysis -> used for checking data
data_raw$CuName <- NULL #Customer can be identified by its identification number, name not used
data_raw$PM <- NULL #PM and AM can be identified by its User ID, name not used
data_raw$AM <- NULL


###############################################################################
# Plausibility of Variables & Identify missing values & eliminate
# =============================================================================
############################################################################### 

#write new data frame for missing value analysis
data_cl <-data_raw

# "#" is a missing value
# 9999999: is a missing value
# 1111111: is not a missing value: Indicates 'no status' in HOM-variables and 'positive FCadj in CostMost-Variables
data_cl[data_cl==9999999] <- NA
data_cl[data_cl == "#"] <- NA

#Set Var NULL, da Verdacht auf nicht korrekte Berechnung. Prob: some BA do not have a median, which could not be the case
data_cl$BAImportPr <- NULL #es kann nicht 100% sichergestellt werden, dass diese Variable korrekt berechnet wurde
data_cl$BUImportPr <- NULL #es kann nicht 100% sichergestellt werden, dass diese Variable korrekt berechnet wurde
data_cl$MSImportPr <- NULL #es kann nicht 100% sichergestellt werden, dass diese Variable korrekt berechnet wurde

#check plausibility of CostMostnegFCajd: Diff btw. Projct Closure Act und last FC adj. muss <= PrtimeAct in months
#if higher, FC has been adjusted after Project Closure (NA have to be excluded of test)
p_monthsnegMS <- data_cl$CostMostnegFCadjMS <= data_cl$PrTimeAct
p_monthsnegME <- data_cl$CostMostnegFCadjME <= data_cl$PrTimeAct
p_monthsnegPA <- data_cl$CostMostnegFCadjPA <= data_cl$PrTimeAct
p_monthsnegIS <- data_cl$CostMostnegFCadjIS <= data_cl$PrTimeAct
p_monthsfirst <- data_cl$CostFirstadj <= data_cl$PrTimeAct

#plausibility and corretness for this group of factors questioned set NULL
data_cl$CostMostnegFCadj <- NULL 
data_cl$CostMostnegFCadjMS <- NULL
data_cl$CostMostnegFCadjME <- NULL
data_cl$CostMostnegFCadjPA <- NULL
data_cl$CostMostnegFCadjIS <- NULL
data_cl$CostFirstadj <- NULL

#count  NA per column
na_count <-sapply(data_cl, function(y) sum(length(which(is.na(y)))))
na_count <- data.frame(na_count)
write.xlsx(na_count, "na_count.xlsx")

#delete columns with NA's abvoe 200
data_cl$PrTimeDelayMS5 <- NULL
data_cl$AMAge2 <- NULL #as age and ten cannot be evaluated, there is little sense to evaluate AM as well
data_cl$AMTen2 <- NULL
data_cl$AMNo <- NULL
data_cl$CostMostnegFCadj <- NULL # erklärung suchen!!!!

# some data loss has to be included otherwise there would be too many columns excluded
# Pr Time Delay and Pr Act are directly connected, therefore both stay in the sample although it has more NA's than others

#remove remaining rows (whole data set) from data
data <- na.omit(data_cl)
sapply(data, function(y) sum(length(which(is.na(y)))))



###############################################################################
# ============================================================================
# Plausibility Tests
# =============================================================================
############################################################################### 

#TOBuD > 0, negative budgets are not plausible
#WA is 4% and missing in data. Bud has 9 cost compponets of which we only evalute 4
#BudMSTot: <= 96, when a MS-Project, >= 0, negative percentages (-CostBudMS/-CostBudTot) are not plausible
#BudMETot: <= 96, when a ME-Project, >= 0, negative percentages (-CostBudME/-CostBudTot) are not plausible
#BudPATot: <= 96, when a PA-Project, >= 0, negative percentages (-CostBudPA/-CostBudTot) are not plausible
#BudISTot: <= 96, when a IS-Project, >= 0, negative percentages (-CostBudIS/-CostBudTot) are not plausible
#DB1Bud: <100 and >0, because projects must have a margin and costs
#DB1Act: <100 because projects must have some cost parts
#SUCostTO <0 is a cost element, (-Cost/TO) is always negative 
#CostActBudRel < -100, if it is assuemd that projects do have costs > 0 (-0 - (-100))/(-100) = -100
#CostActBudMSRel < - 100, if it is assuemd that projects do have costs > 0
#CostActBudMERel < -100, if it is assuemd that projects do have costs > 0
#CostActBudPARel < -100, if it is assuemd that projects do have costs > 0
#CostActBudISRel < -100, if it is assuemd that projects do have costs > 0
#CostMostnegFCadjMS >=0, values are rounded and date FC ajd. would lie after date of Project Closure if values are <0
#CostMostnegFCadjME >=0, values are rounded and date FC ajd. would lie after date of Project Closure if values are <0
#HOMYellCost >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#HOMYellQual >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#HOMYellTime >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#HOMRedCost >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#HOMRedQual >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#HOMRedTime >=0, values <0 imply negative PrTimeBase or that Status ajd. lies before PrStartDate
#BAImportPr >0, because TO values cannot be negative negative values are not plausible
#BUImportPr >0, because TO values cannot be negative negative values are not plausible
#MSImportPr >0, because TO values cannot be negative negative values are not plausible
#PrTimeBase >0
#PrTimeAct > 0
#NoPM >0, on minimum one PM
#AMTen >=0, rounded values could be 0
#PMTen >=0, rounded values could be 0
#NoLeadSASFF, on minimum one
#NoSupplSAS, NoSupplSASMS, NoSupplSASME, NoSupplSASPA, NoSupplSASIS >= 0, 0 if third or own supply
#NoContr >0, on minimum one contract
#DB1Budabs > 0, projects must have costs



#subset data according plausibility indication above
data <- subset(data, TOBud !=0 & 
                 BudMSTot <= 96 & BudMETot <= 96 & BudPATot <= 96 & BudISTot <= 96 &
                 BudMSTot >= 0 & BudMETot >= 0 & BudPATot >= 0 & BudISTot >= 0 &
                 DB1Bud >0 & DB1Bud <100 &
                 DB1Act <100 &
                 CostActBudRel > -100 & CostActBudMSRel > -100 & CostActBudMERel > -100 &
                 CostActBudPARel > -100 & CostActBudISRel > -100 &
                 HOMYellCost >=0 & HOMYellQual >=0 & HOMYellTime >=0 &
                 HOMRedCost >=0 & HOMRedQual >=0 & HOMRedTime >=0 &
                 PrTimeBase >0 & PrTimeAct > 0 &
                 NoPM >0 &
                 PMAge2 >0 & PMTen2 >=0 &
                 NoLeadSASFF >0 &
                 NoSupplSAS >=0 & NoSupplSASMS >=0 &  NoSupplSASME >=0 & NoSupplSASPA >=0 & NoSupplSASIS >=0 &
                 NoContr >0 &
                 DB1Budabs >0)
                  

###############################################################################
# ============================================================================
# Outlier Analysis
# =============================================================================
############################################################################### 

#separate all numeric variables
nums <- sapply(data, is.numeric)
outlier <- data[,nums]

#do summary for data
s_outlier <- summary(outlier)
capture.output(s_outlier, file = "s_outlier.txt")
names <- colnames(outlier)

#calculate items for outlier analysis
min <- round(as.numeric(lapply(outlier, min)),2)
Q1 <- round(as.numeric(lapply(outlier, quantile, probs = 0.25)),2)
means <- round(as.numeric(colMeans(outlier),2))
median <- round(as.numeric(lapply(outlier, quantile, probs = 0.5)),2)
Q3 <- round(as.numeric(lapply(outlier, quantile, probs = 0.75)),2)
max <- round(as.numeric(lapply(outlier, max)),2)
iqr <- as.numeric(Q3-Q1)
iqr1_5_min <- as.numeric(Q1 - 1.5*iqr)
iqr1_5_max <- as.numeric(Q3 + 1.5*iqr)
iqr3_min <- as.numeric(Q1 - 3*iqr)
iqr3_max <- as.numeric(Q3 + 3*iqr)
outlier_xlsx <- data.frame(var = colnames(outlier), min, Q1, means, median, Q3, max, iqr,
                           iqr1_5_min, min<iqr1_5_min,
                           iqr1_5_max, max > iqr1_5_max,
                           iqr3_min, min < iqr3_min,
                           iqr3_max, max > iqr3_max)
colnames(outlier_xlsx) <- c("nums","Min","Q1","Mean","Median", "Q3", "Max","iqr","iqr1_5_min", "T_1.5IQR_Min", 
                       "iqr1_5_max","T_1.5IQR_Max",
                       "iqr3_min","T_3IQR_Min",
                       "iqr3_max","T_3IQR_Max")

#write xlsx for outlier
write.xlsx(outlier_xlsx, "outlier.xlsx")
write.xlsx(data, "02_data_Ro.xlsx")

#exclude outliers
d <- subset(data, DB1Act >= -18.67 & DB1Act <=78.35 &
              CostActBudPARel < 664259.5 & CostActBudISRel < 902779.17)

    