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
# Add further Variables
# =============================================================================
###############################################################################



Success <- d$DB1BudDev>=0
Loss <- d$DB1Act >= 0
Dummy_Success <- as.numeric(d$DB1BudDev>=0)
Dummy_Fail <- as.numeric(d$DB1BudDev<0)
Success_Ampel <- cut(d$DB1BudDev, c(min(d$DB1BudDev), -10, -4, max(d$DB1BudDev)),
                     labels = c("red", "yellow", "green"), include.lowest = T)
Dummy_green <- as.numeric(d$DB1BudDev > -4)
Dummy_yell <- as.numeric(d$DB1BudDev <= -4 & d$DB1BudDev > -10)
Dummy_red <- as.numeric(d$DB1BudDev <= -10)
Delay <- d$PrTimeDelay<0 #no difference if i use Ample System
TOBud_Cat <- as.character(cut(d$TOBud, c(min(d$TOBud), seq(500,5000, 500), 10000, max(d$TOBud)+1), right = F))
TOBudDevabs <- as.numeric(d$TOAct-d$TOBud)
DB1BudDevabs <- as.numeric(d$DB1Actabs-d$DB1Budabs)
CostAct <- as.numeric((-1)*(d$TOAct-d$DB1Actabs))
CostBud <- as.numeric((-1)*(d$TOBud-d$DB1Budabs))
CostBudDevabs <- as.numeric(CostAct-CostBud)
d$PrStartDate <- as.Date(d$PrStartDate, format = "%d.%m.%Y")
Cat_age <- cut(d$PMAge2, seq(20,65,5))

#add variables to data d
d <- data.frame(d, Success = Success, Dummy_Success, Dummy_Fail,Success_Ampel,
                Dummy_green, Dummy_yell, Dummy_red, Loss,
                Delay, TOBud_Cat, TOBudDevabs, DB1BudDevabs, CostBudDevabs, CostAct, CostBud, Cat_age,
                stringsAsFactors = F)

#format all logical und factors data as character
d$Success <- as.character(d$Success)
d$Success_Ampel <- as.character(d$Success_Ampel)
d$Delay <- as.character(d$Delay)
d$Cat_age <- as.character(d$Cat_age)
d$Loss <- as.character(d$Loss)

#write final data to xlsx
write.xlsx(d, "03_d_o.xlsx")


#make two data sets for further anaylzing
uBud <- subset(d, DB1BudDev < 0)
oBud <- subset(d, DB1BudDev >=0)

###############################################################################
# =============================================================================
# 1. Data Overview
# =============================================================================
###############################################################################

#write to excel analysis for furhter analysis
wb = createWorkbook()

#sample overview
ov <- ftable(d$Success)
sheet = createSheet(wb, sheetName = "ov")
addDataFrame(ov, sheet = sheet, row.names = F)



#calculate absolut frequency of Success and Fail pro Region und BA
e_freq <- aggregate(cbind(Dummy_Success, Dummy_Fail) ~ Region + BA + BU + MS + CuNo +
                      PMNo + PMChange + LeadSASPr + LeadSAS.PrFF + CostFCadj + CostFCadjMS+
                      CostFCadjME + CostFCadjPA + CostFCadjIS + NoPM + NoLeadSASFF, d, sum)
Total <- as.numeric(e_freq$Dummy_Fail+e_freq$Dummy_Success)
e_freq <- data.frame(e_freq, Total,
                       Success_Per = e_freq$Dummy_Success/Total, Fail_Per = e_freq$Dummy_Fail/Total,
                       Success_Fail = e_freq$Dummy_Success/e_freq$Dummy_Fail)
sheet = createSheet(wb, sheetName = "ov_success")
addDataFrame(e_freq, sheet = sheet, row.names = F)

#make table for BA/reg frequency
e_freq_reg_BA <- aggregate(cbind(Dummy_Success, Dummy_Fail) ~ Region + BA, data = d, sum)
e_freq_reg_BA <- data.frame(Erfolgsquote = e_freq_reg_BA$Dummy_Success/e_freq_reg_BA$Dummy_Fail,
                           e_freq_reg_BA, 
                           Fail_per = e_freq_reg_BA$Dummy_Fail/(e_freq_reg_BA$Dummy_Fail+e_freq_reg_BA$Dummy_Success),
                           Total = e_freq_reg_BA$Dummy_Success+e_freq_reg_BA$Dummy_Fail,
                           stringsAsFactors = F)
sheet = createSheet(wb, sheetName = "freq_reg_ba")
addDataFrame(e_freq_reg_BA, sheet = sheet, row.names = F)



#make table for BA/BU frequency
e_freq_BA_BU <- aggregate(cbind(Dummy_Success, Dummy_Fail) ~ BA + BU, data = d, sum)
e_freq_BA_BU <- data.frame(Erfolgsquote = e_freq_BA_BU$Dummy_Success/e_freq_BA_BU$Dummy_Fail,
                           e_freq_BA_BU, Total = e_freq_BA_BU$Dummy_Success+e_freq_BA_BU$Dummy_Fail,
                           Fail_per = e_freq_BA_BU$Dummy_Fail/(e_freq_BA_BU$Dummy_Fail+e_freq_BA_BU$Dummy_Success),
                           stringsAsFactors = F)
sheet = createSheet(wb, sheetName = "freq_BA_BU")
addDataFrame(e_freq_BA_BU, sheet = sheet, row.names = F)



#make table for BA frequency
e_freq_BA <- aggregate(cbind(Dummy_Success, Dummy_Fail) ~ BA, data = d, sum)
e_freq_BA <- data.frame(Erfolgsquote = e_freq_BA$Dummy_Success/e_freq_BA$Dummy_Fail,
                        e_freq_BA,
                        Total = e_freq_BA$Dummy_Success+e_freq_BA$Dummy_Fail,
                        Fail_per = e_freq_BA$Dummy_Fail/(e_freq_BA$Dummy_Fail+e_freq_BA$Dummy_Success),
                           stringsAsFactors = F)
sheet = createSheet(wb, sheetName = "freq_BA")
addDataFrame(e_freq_BA, sheet = sheet, row.names = F)



#make table for Region frequency
e_freq_reg <- aggregate(cbind(Dummy_Success, Dummy_Fail) ~ Region, data = d, sum)
e_freq_reg <- data.frame(Erfolgsquote = e_freq_reg$Dummy_Success/e_freq_reg$Dummy_Fail,
                         e_freq_reg,
                         Total = e_freq_reg$Dummy_Success+e_freq_reg$Dummy_Fail,
                         Fail_per = e_freq_reg$Dummy_Fail/(e_freq_reg$Dummy_Fail+e_freq_reg$Dummy_Success),
                         stringsAsFactors = F)
sheet = createSheet(wb, sheetName = "freq_Reg")
addDataFrame(e_freq_reg, sheet = sheet, row.names = F)



#calculate frequency according SUCCESS_AMPEL
e_freq_ampel <- aggregate(cbind(Dummy_green,Dummy_yell, Dummy_red) ~ Region + BA + BU + MS + CuNo, d, sum)
tot_ampe <- as.numeric(e_freq_ampel$Dummy_green + e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red)
e_freq_ampel <- data.frame(e_freq_ampel, Total = tot_ampe,
                     Success_Per = e_freq_ampel$Dummy_green/tot_ampe,
                     Fail_Per = (e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red)/tot_ampe,
                     Success_Fail = e_freq_ampel$Dummy_green/(e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red))

sheet = createSheet(wb, sheetName = "ov_success_ampel")
addDataFrame(e_freq_ampel, sheet = sheet, row.names = F)

###############################################################################
# =============================================================================
# 2. Financial Loss with frame factors
# =============================================================================
###############################################################################

#calculate loss based on SUCCESS incl. all frame factors
e_loss <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                 TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                 CostActBudMSabs, CostActBudMEabs,
                                 CostActBudPAabs, CostActBudISabs) ~ Success + BPMID + Region + BA + BU + MS + CuNo,
                           data = d, sum)
sheet = createSheet(wb, sheetName = "ov_loss")
addDataFrame(e_loss, sheet = sheet, row.names = F)



#calculate loss based on SUCCESS_AMPEL incl. all frame factors
e_loss_ampel <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                CostActBudMSabs, CostActBudMEabs,
                                CostActBudPAabs, CostActBudISabs) ~ Success_Ampel + Region + BA + BU + MS + CuNo,
                          data = d, sum)
sheet = createSheet(wb, sheetName = "ov_loss_ampel")
addDataFrame(e_loss_ampel, sheet = sheet, row.names = F)



#calculate Region BA loss
e_loss_Reg_BA <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                          TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                          CostActBudMSabs, CostActBudMEabs,
                          CostActBudPAabs, CostActBudISabs) ~ Region + BA+Success, data = d, sum)
sheet = createSheet(wb, sheetName = "ov_loss_Reg_BA")
addDataFrame(e_loss_Reg_BA, sheet = sheet, row.names = F)



#calculate Region loss
e_loss_Reg <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                 TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                 CostActBudMSabs, CostActBudMEabs,
                                 CostActBudPAabs, CostActBudISabs) ~ Region +Success, data = d, sum)
sheet = createSheet(wb, sheetName = "ov_loss_Reg")
addDataFrame(e_loss_Reg, sheet = sheet, row.names = F)



#calculate BA loss
e_loss_BA <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                 TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                 CostActBudMSabs, CostActBudMEabs,
                                 CostActBudPAabs, CostActBudISabs) ~ BA+Success, data = d, sum)
sheet = createSheet(wb, sheetName = "ov_loss_BA")
addDataFrame(e_loss_BA, sheet = sheet, row.names = F)

#save to excel
saveWorkbook(wb, "Evaluation.xlsx")

###############################################################################
# =============================================================================
# K.Test
# =============================================================================
###############################################################################

nums_d <- sapply(d, is.numeric)

#transform subsets to data long
uBudl <- melt(uBud[,nums_d], id.vars = "BPMID")
oBudl <- melt(oBud[,nums_d], id.vars = "BPMID")

for(i in 1:length(nums_d)){
  ks <- ks.test(uBud[,nums_d], oBud[,nums_d])
}

###############################################################################
# =============================================================================
# 2. Fulfillment
# =============================================================================
###############################################################################

#create vector for all ff variables which are numeric
ff_num <- c("CostMostnegFCajdMS", "CostMostnegFCajdME", "HOMYellCost", "HOMYellQual", "HOMYellTime",
            "HOMRedCost", "HOMRedTime", "HOMRedQual", "NoPM", "PMAge2", "PMTen2", "NoLeadSASFF")
#create vector for all ff variables which are categorical w/o PMNo as this could give a separate analysis
ff_char <- c("PMChange","LeadSASPr", "LeadSAS.PrFF",
             "CostFCadj", "CostFCadjMS", "CostFCadjME", "CostFCadjPA", "CostFCadjIS")

#create log format of data to iterate through d to crate graphs
dlong <- melt(d, id.vars = c("BPMID","DB1BudDev", "EquLoc", "PMNo", "PMChange",
                             "BA", "BU", "MS", "Region", "CuNo", "PrStartDate",
                             "LeadSASPr","LeadSAS.PrFF", "ConPart",
                             "Success", "Success_Ampel", "Delay", "TOBud_Cat"))

#test combinations of FC depending on success and Region and BA and write to excel
costfcadj <- count(d, c("CostFCadjMS", "CostFCadjME", "CostFCadjPA", "CostFCadjIS","Success", "Region", "BA"))
write.xlsx(costfcadj, "FF-CostFCadjCombi.xlsx")



#HOMStatus
noyellcost <- d$HOMYellCost == 1111111

#histogram for all ff-num
ggplot(subset(dlong, variable == "HOMRedCost"), aes(x = value, fill = Success)) +
  geom_histogram(alpha = 0.5, position = 'identity', bins = 15) +xlim(min(d$HOMRedCost), 300)


#evaluation of PMAge und PMTen
Cat_age <- cut(d$PMAge2, seq(20,65,5))
d <- data.frame(d, Cat_age, stringsAsFactors = F)
Cat_age <- as.character(d$Cat_age)

ffpm <- aggregate(cbind(Dummy_Success, Dummy_Fail)~PMAge2 + PMTen2+Cat_age, data=d, sum)
wb = createWorkbook()
sheet = createSheet(wb, "ffpm")
addDataFrame(ffpm, sheet = sheet)

ffpmchange <- aggregate(cbind(Dummy_Success, Dummy_Fail)~PMChange, data = d, sum)
sheet = createSheet(wb, "ffpmchange")
addDataFrame(ffpmchange, sheet = sheet)

ffpmno <- aggregate(cbind(Dummy_Success, Dummy_Fail)~PMNo, data = d, sum)
sheet = createSheet(wb, "ffpmno")
addDataFrame(ffpmno, sheet = sheet)

ffnopm <- aggregate(cbind(Dummy_Success, Dummy_Fail)~NoPM, data = d, sum)
sheet = createSheet(wb, "ffnopm")
addDataFrame(ffnopm, sheet = sheet)

ffpmageten <- aggregate(cbind(PMAge2, PMTen2)~Success, data = d, mean)
sheet = createSheet(wb, "ffpmageten")
addDataFrame(ffpmageten, sheet = sheet)

ffleadsaspr <- aggregate(cbind(Dummy_Success, Dummy_Fail)~LeadSASPr, data = d, sum)
sheet = createSheet(wb, "ffleadsaspr")
addDataFrame(ffleadsaspr, sheet = sheet)

ffleadsasff <- aggregate(cbind(Dummy_Success, Dummy_Fail)~NoLeadSASFF, data = d, sum)
sheet = createSheet(wb, "ffleadsasff")
addDataFrame(ffleadsasff, sheet = sheet)

ffleadsasprff <- aggregate(cbind(Dummy_Success, Dummy_Fail)~LeadSAS.PrFF, data = d, sum)
sheet = createSheet(wb, "ffleadsasprff")
addDataFrame(ffleadsasprff, sheet = sheet)


saveWorkbook(wb, "ff.xlsx")

###############################################################################
# =============================================================================
# 3. Cost + Erfolgskriteriumg
# =============================================================================
###############################################################################

#make histogram plots for cost factors

cost <- c("TOBud", "BudMSTot", "BudMETot", "BudPATot", "BudISTot",
          "CostBudDevabs","CostActBudMSabs", "CostActBudMEabs", "CostActBudPAabs", "CostActBudISabs",
          "CostActBudRel", "CostActBudMSRel", "CostActBudMERel", "CostActBudPARel", "CostActBudISRel",
          "SUCostTO",
          "DeltaLastFCActMS", "DeltaLastFCActME", "DeltaLastFCActPA", "DeltaLastFCActIS")

#histograms for cost-factors
pdf("cost_histograms.pdf")
plots <- list()
for(i in 1:length(unique(cost))){
plots <- ggplot(subset(dlong, variable == cost[i]), aes(x=value, fill = Success))+
         geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+labs(x = paste(cost[[i]]))
print(plots)
}
dev.off()

#ggplot for TOCat
pdf("cost_hist_TOBud_Cat.pdf")
plot <- ggplot(d, aes(x = TOBud_Cat, fill = Success))+geom_bar(alpha = 0.5, position = 'identity')+
  labs(x = "TOBUD_Cat")
print(plot)
dev.off()

#calculate mean for successful and not successful projects
cost_mean_TOBud_cat <- aggregate(cbind(BudMSTot, BudMETot, BudPATot, BudISTot,
                             CostBudDevabs, CostActBudMSabs, CostActBudMEabs, CostActBudPAabs, CostActBudISabs,
                             CostActBudRel, CostActBudMSRel, CostActBudMERel, CostActBudPARel, CostActBudISRel,
                             SUCostTO)~TOBud_Cat+Success, data = d, mean)

cost_mean_success <- aggregate(cbind(BudMSTot, BudMETot, BudPATot, BudISTot,
                             CostBudDevabs, CostActBudMSabs, CostActBudMEabs, CostActBudPAabs, CostActBudISabs,
                             CostActBudRel, CostActBudMSRel, CostActBudMERel, CostActBudPARel, CostActBudISRel,
                             DeltaLastFCAct, DeltaLastFCActMS, DeltaLastFCActME, DeltaLastFCActPA, DeltaLastFCActIS,
                             SUCostTO)~Success, data = d, mean)

cost_mean_region <- aggregate(cbind(BudMSTot, BudMETot, BudPATot, BudISTot,
                             CostBudDevabs, CostActBudMSabs, CostActBudMEabs, CostActBudPAabs, CostActBudISabs,
                             CostActBudRel, CostActBudMSRel, CostActBudMERel, CostActBudPARel, CostActBudISRel,
                             SUCostTO)~Success+TOBud_Cat+Region, data = d, mean)

cost_sum_TOBudCat <- aggregate(cbind(CostBudDevabs, CostActBudMSabs, CostActBudMEabs,
                             CostActBudPAabs, CostActBudISabs,
                             DeltaLastFCAct, DeltaLastFCActMS, DeltaLastFCActME,
                             DeltaLastFCActIS, DeltaLastFCActPA)~TOBud_Cat+Success+Region, data = d, sum)

#calculate mean of SUCostTO for Success/Fail Projects
cost_su <- aggregate(SUCostTO~Success, data =d, mean)

#calculate mean of SUCostTO for Success/Fail Projects per TOCat
cost_su_TOBudCat <- aggregate(SUCostTO~Success+TOBud_Cat, data =d, mean)


#count freq per TOCat
TOBud_Cat_summary <- count(d, c("Success", "TOBud_Cat"))


wb = createWorkbook()
sheet = createSheet(wb, sheetName = "cost_mean_toBudcat")
addDataFrame(cost_mean_TOBud_cat, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "cost_mean_success")
addDataFrame(cost_mean_success, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "cost_mean_region")
addDataFrame(cost_mean_region, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "cost_sum_tobudcat")
addDataFrame(cost_sum_TOBudCat, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "tobudcat_summary")
addDataFrame(TOBud_Cat_summary, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "cost_su_mean")
addDataFrame(cost_su, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "cost_su_Tobudcat_mean")
addDataFrame(cost_su_TOBudCat, sheet = sheet, row.names = F)


saveWorkbook(wb, "cost.xlsx")

###############################################################################
# =============================================================================
# 4. Time
# =============================================================================
###############################################################################

#calculate average delay per MS and of PrTime
time_mean <- aggregate(cbind(PrTimeBase, PrTimeAct, PrTimeDelay,
                        PrTimeDelayMS2,PrTimeDelayMS8,
                        PrTimeDelayMS10, PrTimeDelayMS11) ~ Success+Dummy_Success+Dummy_Fail, data = d, mean)

#implement binary vairalbs for Delay
DelayMS2 <- as.numeric(d$PrTimeDelayMS2<0)
DelayMS8 <- as.numeric(d$PrTimeDelayMS8<0)
DelayMS10 <- as.numeric(d$PrTimeDelayMS10<0)
DelayMS11 <- as.numeric(d$PrTimeDelayMS11<0)
onTimeMS2 <- as.numeric(d$PrTimeDelayMS2>=0)
onTimeMS8 <- as.numeric(d$PrTimeDelayMS8>=0)
onTimeMS10 <- as.numeric(d$PrTimeDelayMS10>=0)
onTimeMS11 <- as.numeric(d$PrTimeDelayMS11>=0)

#evaluate if a delay is persistent
time_delay <- aggregate(cbind(DelayMS2, onTimeMS2, DelayMS8, onTimeMS8,
                        DelayMS10, onTimeMS10, DelayMS11, onTimeMS11)~Success+Delay, data=d, sum)

#evaluate if sukzessiver delay
time_sukz <- count(d,c("Success","DelayMS2", "DelayMS8", "DelayMS10", "DelayMS11"))

wb = createWorkbook()
p = createSheet(wb, "time_mean")
addDataFrame(time_mean, p)
p = createSheet(wb, "time_delay")
addDataFrame(time_delay, p)
p = createSheet(wb, "time_sukz")
addDataFrame(time_sukz, p)
saveWorkbook(wb, "time.xlsx")

###############################################################################
# =============================================================================
# 5. SQ
# =============================================================================
###############################################################################

#plot histograms for sq-factors
sq <- c("BUORBudGapAbs","BUORBudGapRel","RegiORBudGapAbs", "RegiORBudGapRel")

pdf("sq_histograms.pdf")
plots <- list()
for(i in 1:length(unique(sq))){
  plots <- ggplot(subset(dlong, variable == sq[i]), aes(x=value, fill = Success))+
    geom_histogram(alpha = 0.5, position = 'identity', bins = 15)+labs(x = paste(sq[[i]]))
  print(plots)
}
dev.off()


#calculate mean pressure of ORbudget
sq_mean <- aggregate(cbind(BUORBudGapAbs,BUORBudGapRel,RegiORBudGapAbs, RegiORBudGapRel)~Success,
                data = d, mean)

sq_mean_reg_ba <- aggregate(cbind(BUORBudGapAbs,BUORBudGapRel,RegiORBudGapAbs, RegiORBudGapRel)~Success+Region+BA,
                     data = d, mean)

wb = createWorkbook()
p = createSheet(wb, sheetName = sq_mean)
addDataFrame(sq_mean, "sq_mean")

write.xlsx(sq_mean,"sq.xlsx")

###############################################################################
# =============================================================================
# 6. Complexity
# =============================================================================
###############################################################################

complex <- c("NoSupplSAS", "NoSupplSASMS","NoSupplSASME","NoSupplSASPA","NoSupplSASIS","NoContr")

pdf("complex.pdf")
plots <- list()
for(i in 1:length(unique(complex))){
  plots <- ggplot(subset(dlong, variable == complex[i]), aes(x=value, fill = Success))+
    geom_bar(alpha = 0.5, position = 'identity')+labs(x = paste(complex[[i]]))
  print(plots)
}
dev.off()

complex <- aggregate(cbind(Dummy_Success,Dummy_Fail)~NoSupplSAS +
                     NoSupplSASMS+NoSupplSASME+NoSupplSASPA+NoSupplSASIS+NoContr+ConPart, data = d, sum)

wb = createWorkbook()
p = createSheet(wb, "con_part")
addDataFrame(complex, p)
saveWorkbook(wb, "complex.xlsx")
