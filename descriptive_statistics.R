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

#calculate absolut frequency of Success and Fail pro Region und BA
e_freq <- aggregate(cbind(Dummy_Fail,Dummy_Success) ~ Region + BA + BU + MS + CuNo, d, sum)
Total <- as.numeric(e_freq$Dummy_Fail+e_freq$Dummy_Success)
e_freq <- data.frame(e_freq, Total,
                       Success_Per = e_freq$Dummy_Success/Total, Fail_Per = e_freq$Dummy_Fail/Total,
                       Success_Fail = e_freq$Dummy_Success/e_freq$Dummy_Fail)


#calculate frequency according SUCCESS_AMPEL
e_freq_ampel <- aggregate(cbind(Dummy_green,Dummy_yell, Dummy_red) ~ Region + BA + BU + MS + CuNo, d, sum)
tot_ampe <- as.numeric(e_freq_ampel$Dummy_green + e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red)
e_freq_ampel <- data.frame(e_freq_ampel, Total = tot_ampe,
                     Success_Per = e_freq_ampel$Dummy_green/tot_ampe,
                     Fail_Per = (e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red)/tot_ampe,
                     Success_Fail = e_freq_ampel$Dummy_green/(e_freq_ampel$Dummy_yell + e_freq_ampel$Dummy_red))



###############################################################################
# =============================================================================
# 2. Financial Loss with frame factors
# =============================================================================
###############################################################################

#calculate loss based on SUCCESS incl. all frame factors
e_loss <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                 TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                 CostActBudMSabs, CostActBudMEabs,
                                 CostActBudPAabs, CostActBudISabs) ~ Success + Region + BA + BU + MS + CuNo,
                           data = d, sum)

#calculate loss based on SUCCESS_AMPEL incl. all frame factors
e_loss_ampel <- aggregate(cbind(TOBud, CostBud, DB1Budabs, TOAct, CostAct, DB1Actabs,
                                TOBudDevabs, CostBudDevabs, DB1BudDevabs,
                                CostActBudMSabs, CostActBudMEabs,
                                CostActBudPAabs, CostActBudISabs) ~ Success_Ampel + Region + BA + BU + MS + CuNo,
                          data = d, sum)


#write to excel analysis for furhter analysis
wb = createWorkbook()
sheet = createSheet(wb, sheetName = "ov_success")
addDataFrame(e_freq, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "ov_success_ampel")
addDataFrame(e_freq_ampel, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "ov_loss")
addDataFrame(e_loss, sheet = sheet, row.names = F)
sheet = createSheet(wb, sheetName = "ov_loss_ampel")
addDataFrame(e_loss_ampel, sheet = sheet, row.names = F)
saveWorkbook(wb, "Evaluation.xlsx")

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
             "CostFCadj", "CostFCadjMS", "CostFCadjME", "CostFCadjPA", "CostFCajdIS")


dlong <- melt(d, id.vars = c("BPMID","DB1BudDev", "EquLoc", "PMNo", "BA", "BU", "MS", "Region", "CUNo"))


