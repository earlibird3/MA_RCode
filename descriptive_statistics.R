# =============================================================================
# University of St.Gallen
# Course: Master Thesis
# Authors: Michele Schoch
# Professor: Dr. Simon Pfister
# Date: 22.08.2017
# =============================================================================

#set working directory
setwd('C:/Users/Michèle/Dropbox/Master/MA/06_MA_Files/MA_RCode')
#setwd('C:/Users/Helena Aebersold/Dropbox/Michele_MA/05_ClusterAnalysis')

#source cleaned data
source("./data_cleaning.R")


###############################################################################
# =============================================================================
# 1. Data Overview
# =============================================================================
###############################################################################




#calculate successful and not successfull projects per Region and BA

frequency <- count(d,c("Success", "Region", "BA"))
aggregate(freq ~ Success + Region + BA, frequency, sum)


#calculate loss projects per Region and BA

aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                CostActBudMSabs, CostActBudMEabs, CostActBudPAabs, CostActBudISabs) ~ Success, data = d, sum)

lossx <- x[,c("TOBud", "CostBud", "DB1Budabs", "TOAct","CostAct","DB1Actabs",
          "CostActBudMSabs", "CostActBudMEabs", "CostActBudPAabs", "CostActBudISabs")]
write.xlsx(lossx,"lossx.xlsx")
sapply(lossx, sum)

y <- uBud

lossy <- y[,c("TOBud", "CostBud", "DB1Budabs", "TOAct","CostAct","DB1Actabs",
              "CostActBudMSabs", "CostActBudMEabs", "CostActBudPAabs", "CostActBudISabs")]
sapply(lossy, sum)



###############################################################################
# Descriptive Data Analysis
###############################################################################

#count if project is above budget
overBud = nrow(data_na[data_na$DB1BudDev > 0, ])

#count if project is on budget
onBud = nrow(data_na[data_na$DB1BudDev == 0, ])

#count if project is below budget
underBud = nrow(data_na[data_na$DB1BudDev < 0, ])

totaldata <- data.frame(overBud, onBud, underBud)

#plot histogram of DB1BudDev
pdf("DB1BudDEv_hist.pdf")
DB1BudDev <- ggplot(data_na, aes (x = DB1BudDev, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 20)+
  geom_vline(xintercept = 0)
print(DB1BudDev)
dev.off()


##############################################################################
# Region/BA Datenquerschnitt
###############################################################################

#initialize vectors of unique values of variables to be analyzed
CuNo <- as.character(unique(data_na$CuNo))
Equloc <- as.character(unique(data_na$EquLoc.x))
Region <- as.character(unique(data_na$Region))
PMNo <- as.character(unique(data_na$PMNo))
PMChange <- as.character((unique(data_na$PMChange)))
AMNo <- as.character(unique(data_na$AMNo))
BA <- as.character(unique(data_na$BA))
BU <- as.character(unique(data_na$BU))
MS <- as.character(unique(data_na$MS))
LeadSASPr <- as.character(unique(data_na$LeadSASPr))  
LeadSAS.PrFF <- as.character(unique(data_na$LeadSAS.PrFF))
ConPar <- as.character(unique(data_na$ConPart))


#count number of projects per region over and under budget

ovreg <- vector("numeric", length = length(Region))
unreg <- vector("numeric", length = length(Region))

for(i in 1:length(Region)){
  ovreg[[i]] <- sum(data_na$Region == Region[i] & data_na$DB1BudDev >=0)
  unreg[[i]] <- sum(data_na$Region == Region[i] & data_na$DB1BudDev <0)
  rRegion <- data.frame(Region, ovreg, unreg, ratio = ovreg/unreg)
}


#count projects over and under budget per BA
BA <-ftable(data_na$BA,data_na$success)
print(xtable(as.matrix(BA)), type = "html", file = "BA.html")

ovba <- vector("numeric", length = length(BA))
unba <- vector("numeric", length = length(BA))
for(i in 1:length(BA)){
  ovba[[i]] <- sum(data_na$BU == BA[i] & data_na$DB1BudDev >=0)
  unba[[i]] <- sum(data_na$BA == BA[i] & data_na$DB1BudDev <0)
  rBA <- data.frame(BA, ovba, unba, ratio = ovba/unba)
}


#count projects over and under budget per BU
ovbu <- vector("numeric", length = length(BU))
unbu <- vector("numeric", length = length(BU))
for(i in 1:length(BU)){
  ovbu[[i]] <- sum(data_na$BU == BU[i] & data_na$DB1BudDev >=0)
  unbu[[i]] <- sum(data_na$BU == BU[i] & data_na$DB1BudDev<0)
  rBU <- data.frame(BU, ovbu, unbu, ratio = ovbu/unbu)
}


#count projects over and under budget per MS
ovms <- vector("numeric", length = length(MS))
unms <- vector("numeric", length = length(MS))

for(i in 1:length(MS)){
  ovms[[i]] <- sum(data_na$MS == MS[i] & data_na$DB1BudDev >= 0)
  unms[[i]] <- sum(data_na$MS == MS[i] & data_na$DB1BudDev <0)
  rMS <- data.frame (MS, ovms, unms, ratio = ovms/unms)
}


#
#Complexity-Analysis
#------------------------------------------------------------------------------

#Number of Contracts
NoContr <- ftable(data_na$NoContr, data_na$success)
print(xtable(as.matrix(NoContr)), type = "html", file = "NoContr.html")
file.show("NoContr.html")

#Number of SAS
NoSupplSAS <- ftable(data_na$NoSupplSAS, data_na$success)
print(xtable(as.matrix(NoSupplSAS)), type = "html", file = "NoSupplSAS.html")
file.show("NoSupplSAS.html")

NoSupplSASMS <- ftable(data_na$NoSupplSASMS, data_na$success)
print(xtable(as.matrix(NoSupplSASMS)), type = "html", file = "NoSupplSASMS.html")
file.show("NoSupplSASMS.html")

NoSupplSASME <- ftable(data_na$NoSupplSASME, data_na$success)
print(xtable(as.matrix(NoSupplSASME)), type = "html", file = "NoSupplSASME.html")
file.show("NoSupplSASME.html")

NoSupplSASPA <- ftable(data_na$NoSupplSASPA, data_na$success)
print(xtable(as.matrix(NoSupplSASPA)), type = "html", file = "NoSupplSASPA.html")
file.show("NoSupplSASPA.html")

NoSupplSASIS <- ftable(data_na$NoSupplSASIS, data_na$success)
print(xtable(as.matrix(NoSupplSASIS)), type = "html", file = "NoSupplSASIS.html")
file.show("NoSupplSASIS.html")

#NoSAS: explore combinations
NoSAS <- data_na[, c("NoSupplSAS","NoSupplSASMS", "NoSupplSASME", "NoSupplSASPA", "NoSupplSASIS", "success")]
resultNoSAS <- count(NoSAS, c("NoSupplSAS","NoSupplSASMS", "NoSupplSASME", "NoSupplSASPA", "NoSupplSASIS", "success"))
write.xlsx(resultNoSAS, "NoSAS.xlsx")


#plot histograms for all variables where successful and not successful projects can be viewed
complexity <- c("NoSupplSAS","NoSupplSASMS","NoSupplSASME","NoSupplSASPA","NoSupplSASIS","NoContr" )
complexity <- as.character(complexity)
pdf("hcomplexity.pdf")
histComp <- list()
for (i in 1:length(complexity)){
  histComp <- ggplot(subset(data_nal, variable == complexity[i]), aes(x=value, fill = success))+
    geom_bar(stat = "count", position = "identity", alpha = 0.5) +
    xlab(paste(complexity[i]))
  
  print(histComp)
  
}
dev.off()                            


#plot histograms with Region split for all variables where successful and not successful projects can be viewed
pdf("hcomplexity_region_split.pdf")
histCompReg <- list()
for (i in 1:length(complexity)){
  histCompReg <- ggplot(subset(data_nal, variable == complexity[i]), aes(x=value, fill = success))+
  geom_bar(stat = "count", position = "identity", alpha = 0.5) +
  facet_wrap(~Region) +
  xlab(paste(complexity[i]))

print(histCompReg)

}
dev.off()                            

#
#FF-Analysis
#------------------------------------------------------------------------------
#PM: Analysis of Changes
PMChange <- ftable(data_na$PMChange,data_na$succes)
print(xtable(as.matrix(PMChange)), type = "html", file = "PMChange.html")
file.show("PMChange.html")


#PM: Analysis pf Portfolio of each PM fail projects
PMNo_u <- count(uBud, c("PMNo","success"))
pmnou <- aggregate(list(uBud$TOBud, uBud$TOAct, uBud$TOLossabs, uBud$DB1Budabs, uBud$DB1Actabs, uBud$DB1Lossabs),
          by = list(unique.values = uBud$PMNo), FUN = sum)
colnames(pmnou) = c("PMNo", "TOBud", "TOAct", "TOLoss", "DB1Budabs", "DB1Actabs", "DB1Lossabs")
pmnou <- merge(PMNo_u, pmnou, by = c("PMNo"))
#PM: Analysis pf Portfolio of each PM successful projects
PMNo_o <- count(oBud, c("PMNo","success"))
pmnoo <- aggregate(list(oBud$TOBud, oBud$TOAct, oBud$TOLossabs, oBud$DB1Budabs, oBud$DB1Actabs, oBud$DB1Lossabs),
                   by = list(unique.values = oBud$PMNo), FUN = sum)
colnames(pmnoo) = c("PMNo", "TOBud", "TOAct", "TOLoss", "DB1Budabs", "DB1Actabs", "DB1Lossabs")
pmnoo <- merge(PMNo_o, pmnoo, by = c("PMNo"))
#PM: Analysis pf Portfolio of each PM fail projects
PMNo <- count(data_na, c("PMNo","success"))

wb = createWorkbook()
sheet = createSheet(wb, sheetName = "PMNo")
addDataFrame(PMNo, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "PMNo_fail")
addDataFrame(pmnou, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "PMNo_successful")
addDataFrame(pmnoo, sheet = sheet, row.names = F)
saveWorkbook(wb, "PM_Portfolio.xlsx")


#find PM with more than one project
PMNo_1 <- count(data_na$PMNo)
PMNo_1v <- subset(PMNo_1, freq >5)
PMNo_plots <- as.character(unique(PMNo_1v$x))

#PM: Analysis of Time usage for PM
cols <- c("TRUE"= "#01b558", "FALSE" = "#f04546")
pdf("plots_PM_Time_Analysis.pdf")
plots_time_PM = ls()
for(i in 1:length(PMNo_plots)){
  plots_time_PM <- ggplot(subset(data_na, PMNo == PMNo_plots[i])) +
    geom_segment(aes(x=PrStartDate, xend=PrEndDate, y=TOBud, yend=TOBud, colour = success), size=3)+
    scale_color_manual(values=cols)+
    xlim(as.Date(c("2007-09-01", "2016-02-01")))+
    ylim(100,20000)
  print(plots_time_PM)
}
dev.off()

#PM plot top scorer 10 and 10 low scorer
top_PM <- s

#example of one PM
ggplot(subset(data_na, PMNo == "18628"), aes(colour=success)) +
  geom_segment(aes(x=PrStartDate, xend=PrEndDate, y=TOBud, yend=TOBud), size=3)+
  scale_color_manual(values=cols)+
  xlim(as.Date(c("2007-09-01", "2016-02-01")))+
  ylim(100,35000)



#PM: Analyze number of PM during project lifetime
NoPM <- ftable(data_na$NoPM,data_na$succes)
print(xtable(as.matrix(NoPM)), type = "html", file = "NoPM.html")
#plot histogram
pdf("NoPM.pdf")
NoPMplot <- ggplot(subset(data_nal, variable == "NoPM"), aes(x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha = 0.5)
print(NoPMplot)
dev.off()


#PM: Analyze age of PM
PMAge <- ftable(data_na$PMAge,data_na$succes)
print(xtable(as.matrix(PMAge)), type = "html", file = "PMAge.html")
file.show("PMAge.html")
#plot histogram
ggplot(subset(data_nal, variable == "PMAge"), aes( x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)


#PM: Analyze Tenure
PMTen <- ftable(data_na$PMTen,data_na$succes)
print(xtable(as.matrix(PMTen)), type = "html", file = "PMTen.html")
file.show("PMTen.html")
#plot histogram
ggplot(subset(data_nal, variable == "PMTen"), aes( x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)



#Lead SASPr: count projects over/under budget and ratio for LeadSASPr
LeadSAS <- ftable(data_na$LeadSASPr, data_na$success)
print(xtable(as.matrix(LeadSAS)), type = "html", file = "LeadSAS.html")
file.show("LeadSAS.html")

#Lead SASFF: analyze number of LeadSASF
pdf("NoLeadSASFF.pdf")
NoLeadSASFF <- ggplot(subset(data_nal, variable == "NoLeadSASFF"), aes( x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)
print(NoLeadSASFF)
dev.off()

#Lead SASPrFF: analyze number if LeadSASPr and LeadSASFF are different
pdf("LeadSAS_PRFF.pdf")
LeadSAS_PrFF <- ggplot(data_na, aes( x = LeadSAS.PrFF, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)
print(LeadSAS_PrFF)
dev.off()



#CostFCadj: plot frequency
FF <- c("CostFCadj","CostFCadjMS","CostFCadjME","CostFCadjPA","CostFCadjIS")
FF <- as.character(FF)

pdf("hFF.pdf")
histFF <- list()
for (i in 1:length(FF)){
  histFF <- ggplot(subset(data_nal, variable == FF[i]), aes(x= value, fill = success))+
    geom_bar(stat = "count", position = "identity", alpha = 0.5) +
    xlab(paste(FF[i]))
  print(histFF)
}
dev.off()  

#CostFCajd: evalute several combinations of FCadjustment for unsuccessfull projects
dataID <- seq.int(nrow(data_na))
CF <- data_na[, c("CostFCadjMS", "CostFCadjME", "CostFCadjPA", "CostFCadjIS", "success")]
CF[CF == 0] <- "N"
CF[CF == 1] <- "D"
CF[CF == 2] <- "U"
CF <- data.frame(dataID, CF, stringsAsFactors = F)

oCF <- subset(CF, success == TRUE)
oCFresult <- count(oCF[,-1])
uCF <- subset(CF, success == FALSE)
uCFresult <- count(uCF[,-1])

wb = createWorkbook()
sheet = createSheet(wb, "CF_Success")
addDataFrame(oCFresult, sheet = sheet, startColumn = 1, row.names = F)

sheet = createSheet(wb, "CF_Fail")
addDataFrame(uCFresult, sheet = sheet, startColumn = 1, row.names = F)

saveWorkbook(wb, "CF_Combinations.xlsx")


#mostNegFC vs. Project Closure: plot relation in scatterplott of data subsets Success = T/F
ggplot(data_na, aes(x = MMS_FCadj_PrAct, y= DB1BudDev, colour = success)) + geom_point() + geom_hline(yintercept = 0)+
  geom_vline(xintercept = 1)
ggplot(data_na, aes(x = MME_FCadj_PrAct, y= DB1BudDev, colour = success)) + geom_point() + geom_hline(yintercept = 0)+
  geom_vline(xintercept = 1)

ggplot(data_na, aes(x = MMS_FCadj_PrAct, y= DB1BudDev, colour = success)) + geom_point() + geom_hline(yintercept = 0)+
  geom_vline(xintercept = 0.5)  + xlim(0,1.2) + geom_smooth(method = "lm")
ggplot(data_na, aes(x = MME_FCadj_PrAct, y= DB1BudDev, colour = success)) + geom_point() + geom_hline(yintercept = 0)+
  geom_vline(xintercept = 0.5)  + xlim(0,1.2)

#CONCL: plausibility of these data has to be highly questioned as intuitively 
#CONCL:the number of months from FC and PrEndAct should be lower than the whole PrTimeAct (which is not the case --> graphs, mean)
#CONCL: absolute number of months does not say a lot as Project length differ so 8 months could be at the end or beginging 
#CONCL: of a project if the PrAct is 10 or 48.

#HOM: PrStartDate

ggplot(data_na, aes (x = PrStartDatem, y = DB1BudDevabs, colour = success)) + geom_point(position = position_dodge(1))
ggplot(data_na, aes (x = PrStartDatem, fill = success)) +
  geom_bar(stat ="count", position = "identity", alpha = 0.5)

ggplot(data_na, aes (x = PrStartDateD, fill = success)) +
  geom_bar(stat ="count", position = "identity", alpha = 0.5)

#HOM: Month(Stauts - HOM)/PrBaseline: values abvoe 1 indicate(sample might contain few data which are not plausible max 10)

HOM <- c("HOMYellCost","HOMYellQual","HOMYellTime","HOMRedCost","HOMRedQual","HOMRedTime")
HOM <- as.character(HOM)

pdf("graphs_HOM.pdf")
graphs_HOM <- list()
for (i in 1:length(HOM)){
 graphs_HOM <-  ggplot(subset(data_nal, variable == HOM[i]), aes(x = value, y= DB1BudDev)) +
    geom_point(aes(colour = success)) + xlim(0,2500)+ xlab(paste(HOM[i]))
 print(graphs_HOM)
}
dev.off()

pdf("graphs_HOM_25.pdf")
graphs_HOM_25 <- list()
for (i in 1:length(HOM)){
  graphs_HOM_25 <-  ggplot(subset(data_nal, variable == HOM[i]), aes(x = value, y= DB1BudDev)) +
    geom_point(aes(colour = success)) + xlim(0,25) + xlab(paste(HOM[i]))
  print(graphs_HOM_25)
}
dev.off()

pdf("graphs_HOM_50.pdf")
graphs_HOM_50 <- list()
for (i in 1:length(HOM)){
  graphs_HOM_50 <-  ggplot(subset(data_nal, variable == HOM[i]), aes(x = value, y= DB1BudDev)) +
    geom_point(aes(colour = success)) + xlim(25,50) + xlab(paste(HOM[i]))
  print(graphs_HOM_50)
}
dev.off()

pdf("graphs_HOM_75.pdf")
graphs_HOM_75 <- list()
for (i in 1:length(HOM)){
  graphs_HOM_75 <-  ggplot(subset(data_nal, variable == HOM[i]), aes(x = value, y= DB1BudDev)) +
    geom_point(aes(colour = success)) + xlim(50,75) + xlab(paste(HOM[i]))
  print(graphs_HOM_75)
}
dev.off()

pdf("graphs_HOM_100.pdf")
graphs_HOM_100 <- list()
for (i in 1:length(HOM)){
  graphs_HOM_100 <-  ggplot(subset(data_nal, variable == HOM[i]), aes(x = value, y= DB1BudDev)) +
    geom_point(aes(colour = success)) + xlim(75,100) + xlab(paste(HOM[i]))
  print(graphs_HOM_100)
}
dev.off()

#which status was first? it could be that form the red some reduce to yellow - how many?
Y_C_111 <- data_na$HOMYellCost != 1111111
R_C_111 <- data_na$HOMRedCost != 1111111
Y_T_111 <- data_na$HOMYellCost != 1111111
R_T_111 <- data_na$HOMRedCost != 1111111
Y_R_Cost <- data_na$HOMYellCost<data_na$HOMRedCost
Y_R_Time <- data_na$HOMYellTime<data_na$HOMRedTime
R_Y_Cost <- data_na$HOMYellCost<data_na$HOMRedCost
R_Y_Time <- data_na$HOMYellTime<data_na$HOMRedTime
HOM <- data.frame(Y_C_111, R_C_111, Y_T_111, R_T_111, H_R_Cost, H_R_Time, R_Y_Cost, R_Y_Time, stringsAsFactors = F)
write.xlsx(HOM, "HOM.xlsx", row.names = F)

#
#SQ-Analysis
#------------------------------------------------------------------------------

#Age of AM
ggplot(subset(data_nal, variable == "AMAge"), aes( x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)

#caculated frequency of AMAge and Success
AMAge_Success <- ftable(data_na$AMAge, data_na$success)
print(xtable(as.matrix(AMAge_Success)), type = "html", file = "AMAge.html")
file.show("AMAge.html")

#Tenure of AM
ggplot(subset(data_nal, variable == "AMTen"), aes( x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha=0.5)

#caculated frequency of AMTen and Success
AMTen_Success <- ftable(data_na$AMTen, data_na$success)
print(xtable(as.matrix(AMTen_Success)), type = "html", file = "AMTen.html")
file.show("AMTen.html")

#Budget GAP OR, is there a pressure to sell projects
#histogram plots for BU
ggplot(data_na, aes(x = BUORBudGapAbs, fill = success)) + geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+
  geom_vline(xintercept = 0) + facet_wrap(~BA)

ggplot(data_na, aes(x = BUORBudGapRel, fill = success)) + geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+
  geom_vline(xintercept = 0)
#histogram plots for Region
ggplot(data_na, aes(x = RegiORBudGapAbs, fill = success)) + geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+
  geom_vline(xintercept = 0) + facet_wrap(~Region)
ggplot(data_na, aes(x = RegiORBudGapRel, fill = success)) + geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+
  geom_vline(xintercept = 0)


#calculated mean per BU
OR_Gap_BU_abs <- data.frame(aggregate(data_na$BUORBudGapAbs, by = list(data_na$BU), FUN=median), stringsAsFactors = F)
OR_Gap_BU_rel <- data.frame(aggregate(data_na$BUORBudGapRel, by = list(data_na$BU), FUN=median), stringsAsFactors = F)
#calculated mean per Region
OR_Gap_Reg_abs <- data.frame(aggregate(data_na$RegiORBudGapAbs, by = list(data_na$Region), FUN=median), stringsAsFactors = F)
OR_Gap_Reg_rel <- data.frame(aggregate(data_na$RegiORBudGapRel, by = list(data_na$Region), FUN=median), stringsAsFactors = F)

wb = createWorkbook()

#create sheets
sheet = createSheet(wb, "OR_BU_abs")
addDataFrame(OR_Gap_BU_abs, sheet = sheet, startColumn = 1, row.names = T)
sheet = createSheet(wb, "OR_BU_rel")
addDataFrame(OR_Gap_BU_rel, sheet = sheet, startColumn = 1, row.names = T)
sheet = createSheet(wb, "OR_Reg_abs")
addDataFrame(OR_Gap_Reg_abs, sheet = sheet, startColumn = 1, row.names = T)
sheet = createSheet(wb, "OR_Reg_rel")
addDataFrame(OR_Gap_Reg_rel, sheet = sheet, startColumn = 1, row.names = T)

saveWorkbook(wb, "OR_Gap.xlsx")



#
#Time-Analysis
#------------------------------------------------------------------------------

#plot baseline against act
ggplot(data_na, aes (x = PrTimeBase, y = PrTimeAct)) + geom_point() + geom_abline(slope = 1) + geom_vline(xintercept = c(5,10,15,20))

#Cross Table onTime and Success
CrossTable(data_na$onTime, data_na$success)
onTime_Succes <- table(data_na$onTime, data_na$success)
chisq.test(onTime_Succes)

#Calculate loss per Time category for unsuccessful projects
delay_u <- count(uBud, c("PrTimeDelay_cat","success"))
delay_agg_u <- aggregate(list(uBud$TOBud, uBud$DB1Budabs, uBud$DB1Actabs, uBud$DB1Lossabs),
                 by = list(unique.values = uBud$PrTimeDelay_cat), FUN = sum)
colnames(delay_agg_u) = c("PrTimeDelay_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")
delay_agg_u <- merge(delay_u, delay_agg_u, by = c("PrTimeDelay_cat"))

#Calculate loss per Time category for successful projects
delay_o <- count(oBud, c("PrTimeDelay_cat","success"))
delay_agg_o <- aggregate(list(oBud$TOBud, oBud$DB1Budabs, oBud$DB1Actabs, oBud$DB1Lossabs),
                       by = list(unique.values = oBud$PrTimeDelay_cat), FUN = sum)
colnames(delay_agg_o) = c("PrTimeDelay_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")
delay_agg_o <- merge(delay_o, delay_agg_o, by = c("PrTimeDelay_cat"))

#Calculate net loss per Time category for all projects
delay_agg <- aggregate(list(data_na$TOBud, data_na$DB1Budabs, data_na$DB1Actabs, data_na$DB1Lossabs),
                         by = list(unique.values = data_na$PrTimeDelay_cat), FUN = sum)
colnames(delay_agg) = c("PrTimeDelay_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")

wb = createWorkbook()
sheet = createSheet(wb, sheetName = "Delay_net")
addDataFrame(delay_agg, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "Delay_fail")
addDataFrame(delay_agg_u, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "Delay_successful")
addDataFrame(delay_agg_o, sheet = sheet, row.names = F)
saveWorkbook(wb, "Delay.xlsx")


      
      
#plot PrTimeDelay against DB1BudDEv
ggplot(oBud, aes(x = DB1BudDevabs, y = PrTimeDelay)) + geom_point() + geom_hline(yintercept = 0)
ggplot(uBud, aes(x = DB1BudDevabs, y = PrTimeDelay)) + geom_point() + geom_hline(yintercept = 0)

#at which stage is the delay growing, make 9 categories
ggplot(data_na, aes(x=PrTimeDelayMS2, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=PrTimeDelayMS8, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=PrTimeDelayMS10, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=PrTimeDelayMS11, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)

  
#
#COST-Analysis
#------------------------------------------------------------------------------
#Calculate loss per TO category for unsuccessful projects
loss_u <- count(uBud, c("TOBud_cat","success"))
loss_agg_u <- aggregate(list(uBud$TOBud, uBud$DB1Budabs, uBud$DB1Actabs, uBud$DB1Lossabs),
                         by = list(unique.values = uBud$TOBud_cat), FUN = sum)
colnames(loss_agg_u) = c("TOBud_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")
loss_agg_u <- merge(loss_u, loss_agg_u, by = c("TOBud_cat"))

#Calculate loss per TO category for successful projects
loss_o <- count(oBud, c("TOBud_cat","success"))
loss_agg_o <- aggregate(list(oBud$TOBud, oBud$DB1Budabs, oBud$DB1Actabs, oBud$DB1Lossabs),
                        by = list(unique.values = oBud$TOBud_cat), FUN = sum)
colnames(loss_agg_o) = c("TOBud_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")
loss_agg_o <- merge(loss_o, loss_agg_o, by = c("TOBud_cat"))

#Calculate net loss per TO category for all projects
loss_agg <- aggregate(list(data_na$TOBud, data_na$DB1Budabs, data_na$DB1Actabs, data_na$DB1Lossabs),
                        by = list(unique.values = data_na$TOBud_cat), FUN = sum)
colnames(loss_agg) = c("TOBud_cat" , "TOBud", "DB1Budabs", "DB1Actabs","DB1Lossabs")


wb = createWorkbook()
sheet = createSheet(wb, sheetName = "Loss_net_TO_cat")
addDataFrame(loss_agg, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "Loss_fail_TO_cat")
addDataFrame(loss_agg_u, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "Loss_successful_TO_cat")
addDataFrame(loss_agg_o, sheet = sheet, row.names = F)
saveWorkbook(wb, "Loss_DB1_TO_cat.xlsx")

#TOCat: analyze Time Delays for failures
write.xlsx(TOBud_Delay <- count(uBud, c("TOBud_cat", "PrTimeDelay_cat","success")), "TOBud_Delay.xlsx")

#plot TOBud as histogram absolut
pdf("Cost_TOBud_abs.pdf")
Cost_TOBud_abs <- ggplot(data_na, aes(x = TOBud, fill = success)) +
  geom_histogram(alpha = 0.5, position = "identity", bins = 15)
print(Cost_TOBud_abs)
dev.off()


#plot TOBud per TO Category
pdf("Cost_TOBud_abs_class.pdf")
Cost_TOBud_abs_class <- ggplot(data_na, aes(x = TOBud_cat, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha = 0.5)
print(Cost_TOBud_abs_class)
dev.off()

ggplot(subset(data_nal, variable == "TOBud"), aes(x = value, fill = success)) +
  geom_bar(stat = "count", position = "identity", alpha = 0.5)

#Bud parts
ggplot(data_na, aes(x=BudMSTot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=BudMSTot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) + xlim(70,90)
ggplot(data_na, aes(x=BudMETot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=BudMETot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) + xlim(0, 20)
ggplot(data_na, aes(x=BudPATot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)
ggplot(data_na, aes(x=BudPATot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) + xlim(21, 100)
ggplot(data_na, aes(x=BudISTot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) 
ggplot(data_na, aes(x=BudISTot, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) + xlim(0, 10)

#DB1BudDev: plot distribution DB1BudDev
pdf("DB1BudDev_Histogram.pdf")
DB1BudDev <- ggplot(data_na, aes(x=DB1BudDev)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15) + xlim (-50, 50)
print(DB1BudDev)
dev.off()


#DB1 Bud: check success rate of projechts with DB1Bud around critical treshold of 23%
sum(data_na$DB1BudDev >=0 & data_na$DB1Bud >=22 & data_na$DB1Bud <=23)
sum(data_na$DB1BudDev <0 & data_na$DB1Bud >=22 & data_na$DB1Bud <=23)
sum(data_na$DB1BudDev >=0 & data_na$DB1Bud <22)
sum(data_na$DB1BudDev <0 & data_na$DB1Bud <22)
sum(data_na$DB1BudDev >=0 & data_na$DB1Bud >23)
sum(data_na$DB1BudDev <0 & data_na$DB1Bud >23)

#DB1Act
summary(oBud$DB1Act)
summary(uBud$DB1Act) # on outlier did not influence mean and median significantly
#eleminate outlier
uBud_out <- subset(uBud, DB1Act>-200)
summary(uBud_out$DB1Act)

#Cost overrun abs
ggplot(data_na, aes(x=CostActBudMSabs, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-100,100)
ggplot(data_na, aes(x=CostActBudMEabs, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-100,100)
ggplot(data_na, aes(x=CostActBudPAabs, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50,50)
ggplot(data_na, aes(x=CostActBudISabs, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50,50)

#Cost overrun rel
ggplot(data_na, aes(x=CostActBudMSRel, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50, 50)
ggplot(data_na, aes(x=CostActBudMERel, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50,50)
ggplot(data_na, aes(x=CostActBudPARel, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50,50)
ggplot(data_na, aes(x=CostActBudISRel, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 15)+
  xlim(-50,50)

#SUCostTO
sum(data_na$SUCostTO == 0)
ggplot(subset(data_na, SUCostTO < 0), aes(x=SUCostTO, fill = success)) +
  geom_histogram(alpha = 0.5, position = "identity", bins = 15) +xlim (-1,-10)
ggplot(subset(data_na, SUCostTO == 0), aes(x=SUCostTO, fill = success)) +
  geom_histogram(alpha = 0.5, position = "identity", bins = 15)

#Last Forecast: view files uBud and oBud summary


#
#framework-Analysis
#------------------------------------------------------------------------------

#evaluate frequency of combinations
Rahm_all <- count(data_na, c("CuNo","Region","EquLoc.x", "BA", "BU","MS","LeadSASPr","success")) #maxfreq = 2 of unsuccessful projects
write.xlsx(Rahm_all, "Rahmenbed_all.xlsx")

Rahm_all_oEqu <- count(data_na, c("CuNo","Region", "BA", "BU","MS","LeadSASPr","success"))
write.xlsx(Rahm_all_oEqu, "Rahmenbed_all_oEqu.xlsx")

Rahm_all_oEqu_oCuNo <- count(data_na, c("Region", "BA", "BU","MS","LeadSASPr","success"))
write.xlsx(Rahm_all_oEqu_oCuNo, "Rahmenbed_all_oEqu_Cu.xlsx")

Rahm_all_oEqu_oCuNo_oLeadPr <- count(data_na, c("Region", "BA", "BU","MS","success"))
write.xlsx(Rahm_all_oEqu_oCuNo_oLeadPr, "Rahmenbed_all_oEqu_Cu_LeadPr.xlsx")


#evaluate importance of projects for BA, BU, MS
pdf("BAImpor.pdf")
BAImport <- ggplot(data_na, aes(x = BAImportPr, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 20) +
  xlim(0,500) + geom_vline(xintercept = 100)
print(BAImport)
dev.off()

pdf("BUImpor.pdf")
BUImport <- ggplot(data_na, aes(x = BUImportPr, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 20) +
  xlim(0,500) + geom_vline(xintercept = 100)
print(BAImport)
dev.off()

pdf("MSImpor.pdf")
MSImport <- ggplot(data_na, aes(x = MSImportPr, fill = success)) + geom_histogram(alpha = 0.5, position = "identity", bins = 20) +
  xlim(0,500) + geom_vline(xintercept = 100)
print(BAImport)
dev.off()

###############################################################################
# =============================================================================
#Analyse Verlust/Gewinn - allgmeiner overview
# =============================================================================
###############################################################################