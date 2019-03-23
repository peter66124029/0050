###################################################################################################
# Purpose:  1. 將台灣五十成分股股價讀入資料，並計算平均數及變異數，不同期間長度所計算之結果：
#           2. 計算Black Litterman model 中所需要之Pick matrix, 係根據不同之 Momentum strategy 所得
# date range:  20040101~20161231，合計?週(20040101~20161231)，或共同156個月份;
# Output: 
# 台灣五十股票之平均報酬率及變異數，該資料是以週資料排列，每週之平均數或變異數之計算，
# 是根據每週時點回溯過去5年的週報酬率資料所計算出來，例如第一筆資料為20030704，則回溯過去五年(19980704~20030704)
# 共有262週歷史資料報酬率，並計算累計報酬率，依據該累計報酬率，形成momentum strategy; 
# 每週得出50-by-50的共變異數矩陣，因此給你該項資料檔共有19600-by-50 (50*392=19600)的文字檔，
# 因為有二種計算cov的方法，因此有二個檔案。另外平均數檔案為392-by-50的文字檔。 
# factor.model.stat {BurStFin}: to estimate covariance matrix with missing values therein
# var.shrink.eqcor {BurStFin}: Ledoit-Wolf Shrinkage Variance Estimate 
######################################################################################################
    ######################################################################
    # create Black Litterman's views based on momentum strategy
    # Momentum strategy is formed by using the past accumulated returns: 
    # 5wks, 10wks, 15wks, 20wks, 25wks, 50wks, 100wks, 150wks, 200wks, 
    # all historical periods.
    ######################################################################
    # q 代表選取股票累積報酬率前(1-q)百分比的股票;
    # q=0.8，代表前20%的股票;
    #############################################################
# XLConnect: Excel Connector for R (that's what this demo is about)
# fImport: Rmetrics - Economical and Financial Data Import
# forecast: Forecasting functions for time series
# zoo: S3 Infrastructure for Regular and Irregular Time Series
# ggplot2: An implementation of the Grammar of Graphics

#install.packages(c("XLConnect", "fImport", "forecast", "zoo", "ggplot2"))

#################
# Load Packages #
#################
rm(list=ls())
#install.packages("fPortfolio")
library(fPortfolio)
library(reshape2)
library(xts)
library(zoo)
#require(zoo)
require(ggplot2)
#library(xts)
library(timeSeries)
#To install a package from within R, you can issue a command like: 
#install.packages("BurStFin", repos="http://www.burns-stat.com/R") 
#install.packages("BurStFin")
library(BurStFin)
#install.packages("BLCOP")
library(BLCOP)
#install.packages("FRAPO")
library(FRAPO)
library(magrittr)
library(PerformanceAnalytics)
#install.packages("xtable")
library(xtable) #  print output in latex
#install.packages("data.table")
library(data.table) # provides a better way to transpose data frame
#============================================================================
# Reading / Importing histocial daily closing stock prices and market value
# from hist_price_1999_2017.txt#
# in-sample data range starts from 2004-2016
# But sample data range starts 5 year earlier, i.e., 1999-2016
# Keep 5 years data because we may need longer data to compute covariance
#=============================================================================
# Importing data: This part of codes import data and save as .rds files so that we don't have 
# run these codes every time!
#==================================================================
#setwd("D:/Data TWN50")
tw50.code<-read.csv("tw50_hist_code.csv", fileEncoding = "UTF-8-BOM")
tail(tw50.code)
write.table(tw50.code, "tw50_hist_code.txt")
#hist.prices<-read.table("hist_price_2003_2016.txt")
hist.prices<-read.table("hist_price_1999_2017_ansi.txt", header = TRUE)
head(hist.prices)
colnames(hist.prices)<-c("id", "name", "date", "close", "mkt")
prices<-hist.prices[,c(1,3,4)]
mkt.value<-hist.prices[,c(1,3,5)]
rm(hist.prices)
hist.prices.df<-dcast(prices,date~id)
dim(hist.prices.df)
last.row<-dim(hist.prices.df)[1]
hist.prices.df[last.row, 1:5]
# delete last row because it's NA (not trading day for 20161231)
#hist.prices.df<-hist.prices.df[-last.row,]
mkt.value.df<-dcast(mkt.value, date~id)
mkt.value.df<-mkt.value.df[-last.row,]
write.csv(hist.prices.df, file= "hist.prices.csv")
#
etf50<-read.table("tw50_benchmarkt.txt", stringsAsFactors = FALSE)
etf50<-etf50[-1,-c(1,2)]
colnames(etf50)<-c("date", "price")
tail(etf50)
#
saveRDS(etf50, "etf50.rds")
saveRDS(hist.prices.df, "hist_prices_df.rds")
saveRDS(mkt.value.df, "mkt_value_df.rds")
saveRDS(tw50.code, "tw50_code.rds")
rm(hist.prices.df)
rm(mkt.value.df)
rm(tw50.code)
#---------------------------------------------------------
# Importing data ends here
#============================================================
#
# load adjusted data from *.rds files 
#setwd("D:/Data TWN50")
hist.prices.df <- readRDS("hist_prices_df.rds")
# prices start from 1999
dim(hist.prices.df)
mkt.value.df<-readRDS("mkt_value_df.rds")
tw50.code<-readRDS("tw50_code.rds")
etf50.df<-readRDS("etf50.rds")
dim(tw50.code)
#=====================================================================================================
# Missing data adjustment: stocks are adjusted:5880, 4938
# Adjusted data has been saved as *.rds and therefore we don't have to rerun the following codes again
#=====================================================================================================
# Read in financial index returns to impute the missing values for stock id = 5880 (合庫) 
#which is listed on 2011/12/1
fin.idx<-read.table("fin_index_1999_2017.txt")
fin.idx<-fin.idx[ ,c(3,4)]
fin.idx<-fin.idx[-1,]
length(fin.idx)
head(fin.idx)
tail(fin.idx)
#[1] 4746
names(fin.idx)<-c("date", "return")
idx.date<-as.Date(fin.idx$date, format="%Y%m%d")
fin.idx<-xts(as.numeric(as.character(fin.idx[,2])), order.by=idx.date)
fin.idx<-coredata(fin.idx)
t<-hist.prices.df["5880"]
t.1<-t[is.na(t)]
k<-length(t.1)
k
# Here we try to use financial index returns to proxy for 5880 stock returns 
# in order to impute the missing prices
i=3246
for (i in k:1){
  hist.prices.df[i,"5880"]<-as.numeric(hist.prices.df[i+1,"5880"])/(1+(fin.idx[i+1]/100))
}

length(hist.prices.df[,"5880"])
price.5880<-data.frame(as.numeric(hist.prices.df[,"5880"]))
price.5880<-xts(price.5880, order.by=idx.date)
dim(price.5880)
head(price.5880)
tail(price.5880)
colnames(price.5880)<-"price"
length(price.5880)

# Read outstanding shares to impute market value of 5880
shares.5880<-read.table("5880_mkv.txt", stringsAsFactors = FALSE)
shares.5880<-shares.5880[, -c(1,2,5,6)]
shares.5880<-shares.5880[-1,]
names(shares.5880)<-c("date", "shares")
idx.date1<-as.Date(shares.5880$date, format="%Y%m%d")
shares.5880<-xts(as.numeric(shares.5880[,2]), order.by = idx.date1)
colnames(shares.5880)<-"shares"
head(shares.5880)
tail(shares.5880)
#
mkv.5880<-merge(price.5880, shares.5880)
mkv.5880<-na.locf(mkv.5880, fromLast = TRUE)
mkv.5880$mkv<-mkv.5880$price * mkv.5880$shares /1000
head(mkv.5880)
tail(mkv.5880)
#
dim(mkt.value.df)
head(mkt.value.df)
tail(mkt.value.df$date)
mkt.value.df[,"5880"]<-mkv.5880$mkv

#===========================================================================================
# Here we use TEJ (M2300) 電子類指數 returns to impute stock 4938 (和碩) missing prices
# however, M2300 index only starts from 20070703
ele.idx<-read.table("eletr_index_1999_2017_1.txt")
ele.idx<-ele.idx[ ,c(3,4)]
ele.idx<-ele.idx[-1,]
names(ele.idx)<-c("date", "return")
idx.date<-as.Date(ele.idx$date, format="%Y%m%d")
ele.idx<-xts(as.numeric(as.character(ele.idx[,2])), order.by=idx.date)
ele.idx<-coredata(ele.idx)
t<-hist.prices.df["4938"]
t.1<-t[is.na(t)]
k<-length(t.1)
k
for (i in k:1){
  hist.prices.df[i,"4938"]<-as.numeric(hist.prices.df[i+1,"4938"])/(1+(ele.idx[i+1]/100))
}
#--------------------------------------------------------------------------------------
tail(hist.prices.df$date)
saveRDS(hist.prices.df, "hist_prices_df.rds")
saveRDS(mkt.value.df, "mkt_value_df.rds")
# data adjustment ends here
#=======================================================================================

# Here we extract 50 stocks daily prices and market value by month from 200401 to 201612
# The results are saved as tw50.hist.prices[[i]] and tw50.mkt.value[[i]]
# Take out code only 
tw.50.code<-tw50.code[, seq(2,112, by=2)]
m<-length(tw.50.code)
# m: there are 56 quarters from 2004Q1 to 2016Q4
id.i<-tw.50.code[,1]
#extract stock prices
yr.qi<-hist.prices.df[as.character(id.i)]
# convert year-month-day into year-month format, time series;
insample.range<-as.Date(seq(as.yearmon("2004-01-31"),as.yearmon("2017-12-31"),
                            by=1/12), frac=1)
# convert to character format
insample.range<-format(insample.range, "%Y-%m")
# n = the number of months from 2004 to 2017 = 168 months
n<-length(insample.range)
#insample.range[1]
tw50.hist.prices<-list()
tw50.mkt.value<-list()
j=1
i=1
# stock codes will remain the same for three months for each quarter (m=56)
# therefore, i = 1, 2 and 3 when j  = 1; i= 4, 5, and 6 when j = 2, etc.
# i  = 1, 2, ..., 168. i=1 (2004/01), i=2 (2004/02), etc.
for (j in 1:m){
  for (i in (3*j-2):(3*j)){
  tw50.hist.prices[[i]]<-hist.prices.df[as.character(tw.50.code[,j])]
  tw50.hist.prices[[i]]<-as.data.frame(lapply(tw50.hist.prices[[i]], as.numeric))
  date<-as.character(hist.prices.df[,1])
  index<-as.Date(date, format="%Y%m%d")
  tw50.hist.prices[[i]]<-xts(tw50.hist.prices[[i]], order.by = index)
  #tw50.hist.prices[[i]]<-interpNA(tw50.hist.prices[[i]], method = "after")
  tw50.hist.prices[[i]]<-na.locf(tw50.hist.prices[[i]], fromLast = TRUE)
  # extract market value data
  tw50.mkt.value[[i]]<-mkt.value.df[as.character(tw.50.code[,j])]
  tw50.mkt.value[[i]]<-xts(tw50.mkt.value[[i]], order.by = index)
  #tw50.mkt.value[[i]]<-xts(tw50.mkt.value[[i]], order.by = index)
  tw50.mkt.value[[i]]<-tw50.mkt.value[[i]][endpoints(tw50.mkt.value[[i]], on="months"),]
  }
}
# assign names "2004-01"... to price list
names(tw50.hist.prices)<-insample.range
names(tw50.mkt.value)<-insample.range
#=========================================================================================
#==========================================================================
# 均衡資產超額報酬率 Pi = delta * Sigma * w.star
# delta= risk aversion parameter 
# Here we assume an arbitrary value for delta;
# delta = 2
# pi.eq = delta * cov.i %*% mkt.value.m.share 
# #---------------------------------------------------------------------
# # calculate cumulative returns
# cum.period.d = 182 # 6 months  = 180 days
# end.d<-tail(time(window.i),1)
# start.d<-end.d - cum.period.d
# cum.period<-paste(start.d, end.d, sep="/")
# accum.ret<-colSums(window.i[cum.period], na.rm = TRUE)
# # create views matrix = pick
# pick.6mret.ls<-list()
# qq = c(0.7, 0.8, 0.9)
# # assume qq = 0.7
# k=1
# q=qq[k]
# # find the q% cumulative return
# q.6m<-quantile(accum.ret,q)
# qq.6m<-accum.ret>=q.6m
# qq.6m.mtx<-matrix(as.numeric(qq.6m),ncol=50)
# # nw: number of winners (q=80% quantile)
# nw<-sum(qq.6m)
# pick.6m<- apply(qq.6m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=-1/(50-nw))
# i=1
# pick.6mret.ls[[i]]<-as.vector(pick.6m)
# #-------------------------------------------------------------
# # By J.M. Mulvey et al. "Linking Momentum Strategies with
# # Single-Period Portfolio Models", Chapter 18  
# # create confidence level Omega = (1/c -1)*P * Sigma * P'
# # where P = pick matrix, Omega = variance of the view.  
# # The smaller c is, the larger the Omega, the less certain about the views; 
# # The smaller c, the larger the uncertainty!
# c=c(0.01,0.5,0.99)
# v=0.01
# # Calculate confidence level by setting c = 0.01
# pick.6m.mtx<-matrix(pick.6m, nrow=1)
# cov.i.mtx<-matrix(cov.i, nrow=50)
# omega1=(1/c[1]-1)*(pick.6m.mtx%*%cov.i.mtx%*%t(pick.6m.mtx))
# omega2=(1/c[2]-1)*(pick.6m.mtx%*%cov.i.mtx%*%t(pick.6m.mtx))
# omega3=(1/c[3]-1)*(pick.6m.mtx%*%cov.i.mtx%*%t(pick.6m.mtx))
# cl=c(omega1, omega2, omega3)
# # create view matrix
# views <- BLViews(P = pick.6m.mtx, q = v, confidences = cl[1], 
#                  assetNames = colnames(cov.i))
# 
# # Pi.eq = equilibrium returns for 50 stocks
# Pi.eq <-as.numeric(Pi)
# priorVarcov <- cov.i.mtx
# # Assume tau  = 0.5
# marketPosterior <- posteriorEst(views = views, sigma = priorVarcov, mu = pi.eq, tau = 1/2)
# str(marketPosterior)
#========================================================
# Extract date range
# get index first
yr.mon<-unique(format(index, format="%Y-%m")) 
#yr.mon<-as.Date(yr.mon, format = "%Y-%m")
#insample.start.mon<-which(yr.mon=="2003-01")
insample.start.mon<-which(yr.mon=="2003-01")
insample.start.mon
insample.end.mon<-which(yr.mon=="2003-12")
insample.end.mon
outsample.end.mon<-which(yr.mon == "2016-12")
outsample.end.mon
sample.n<-outsample.end.mon - insample.end.mon 
sample.mon<-yr.mon[insample.start.mon:outsample.end.mon]
#-----------------------------------------------------------
length(sample.mon)

x<-sample.mon[1:156]
x<-sample.mon[16]
#cum.n = calculate cumulative returns of past n days
cum.n <- 180
# clevel = confidence level
# clevel=c(0.01,0.5,0.99)
clevel = 0.99
# qt = quantile
# qt = c(0.7, 0.8, 0.9)
qt = 0.7
# v: the views on the asset returns
v=0.01
# delta = risk aversion parameter
delta = 1
kappa=1
tau=1
# pi.eq = equilibrium returns for assets
#cum.n = calculate cumulative returns of cum.n days
BL <- function(x, cum.n, clevel, v, qt, tau, kappa){
  start.m<-which(yr.mon == x)
  end.m<-start.m + 11
  insample.range<-paste(yr.mon[start.m], yr.mon[end.m], sep="/")
  #outsample.m = "2004-01", the first list element of the data
  outsample.m<-end.m + 1
  yr_mon<-yr.mon[end.m]
  yr.qi.ts<-tw50.hist.prices[[yr.mon[outsample.m]]]
  yr.qi.ts<-yr.qi.ts[paste("/", yr_mon, sep="")]
  yr.qi.ts.1<-coredata(yr.qi.ts)
  # compute daily returns
  #qi.ret<-diff(log(yr.qi.ts.1))
  qi.ret<-(yr.qi.ts - lag(yr.qi.ts, 1))/yr.qi.ts 
  qi.ret<-qi.ret[-1,]
  #extract market value data list
  yr.qi.mkt.ts<-tw50.mkt.value[[yr.mon[outsample.m]]]
  #yr.qi<-coredata(yr.qi.ts[insample.range])
  #qi.ret<-diff(log(yr.qi))
  #qi.ret<-qi.ret[-1,]
  mkt.value.m.i<-yr.qi.mkt.ts[yr.mon[outsample.m]]
  mkt.value.m.share<-as.numeric(mkt.value.m.i)/sum(as.numeric(mkt.value.m.i))
  mkt.value.m.share<-as.matrix(mkt.value.m.share, ncol=1)
  # use 250 days return to compute covariance matrix
  Rw<-as.matrix(tail(qi.ret, cum.n))
  #covi<-cov(Rw)
  # Use BurStFin package to compute covariance matrix by allowing short of data length
  covi<-factor.model.stat(Rw)
  delta=1
  pi.eq <- c(delta * covi %*% mkt.value.m.share )
  qi.ret.1<-qi.ret + 1
  cum.ret<-cumprod(tail(qi.ret.1, cum.n))
  cum.ret<-tail(cum.ret,1)
  qt.ret<-quantile(tail(cum.ret, 1), qt)
  q01.ret<-cum.ret>=qt.ret
  q01.ret<-matrix(as.numeric(q01.ret),ncol=50)
  nw<-sum(q01.ret)
  picki<- apply(q01.ret, 2, FUN=function(x) if (x==1) x=1/nw else x=-1/(50-nw))
  picki<- matrix(picki, nrow=1)
  omega=(1/clevel-1)*(picki%*%covi%*%t(picki))
  omega=c(omega)
  views <- BLViews(P = picki, q = v, confidences = omega, assetNames = colnames(qi.ret))
  post <- posteriorEst(views, mu = pi.eq, tau = tau, sigma = covi, kappa = kappa)
  return(post)
}

# PostDist<-list()
# sample.mon[156]
#"2015-12"
length(sample.mon)
# [1] 168
idx<-sample.mon[1:156]
PostDist <- lapply(idx, BL, cum.n = 180, clevel=0.99, v=0.01, qt=0.7, tau = 1, kappa = 1)

## Defining portfolio specifications
EstPrior <- function(x, spec = NULL, ...){
  list(mu = BLR@priorMean, Sigma = BLR@priorCovar)
}

EstBL <- function(x, spec = NULL, ...){
  list(mu = BLR@posteriorMean, Sigma = BLR@posteriorCovar)
}



## Prior specificication
MSPrior <-  portfolioSpec()
setEstimator(MSPrior) <- "EstPrior"
## BL specification
MSBl <-  portfolioSpec()
setEstimator(MSBl) <- "EstBL"
## Constraints
NAssets = 50
#BoxC <- c("minW[1:NAssets] = 0.00", "maxW[1:NAssets] = 1.0")
#BoxC <- c("minW[1:NAssets] = -0.8", "maxW[1:NAssets] = 0.8")
#==============================================================
# Maximum Sharpe ratio portfolio back-test
#==============================================================
## Back test function
## CDaR portfolio 
DDbound <- 0.10 
DDalpha <- 0.95
# PMTD(), minimum tail-dependent portfolio
# PMD(), most diversified portfolio
# PERC(), equal risk contributed portfolios
# PCDaR(), portfolio optimization with conditional draw-down at risk constraint
# PAveDD(), portfolio optimization with average draw-down constraint
# PMaxDD(), portfolio optimization with maximum draw-down constraint
# PMinCDaR(), portfolio optimization for minimum conditional draw-down at risk.
BT <- function(DataSub, BLR, cum.n){
  DataSub <- as.timeSeries(DataSub)
  qi.ret<-(DataSub - lag(DataSub, 1))/DataSub
  qi.ret<-qi.ret[-1,]
  qi.ret<-tail(qi.ret, cum.n)
  PPrior <- minvariancePortfolio(data = qi.ret, spec = MSPrior,
                              constraints = "LongOnly")
  PBl <- minvariancePortfolio(data = qi.ret, spec = MSBl,
                                 constraints = "LongOnly")
  prices<-tail(DataSub, cum.n)
  cd <- PCDaR(prices, alpha = DDalpha, bound = DDbound, 
              softBudget = TRUE)
  wgmv<-Weights(PGMV(qi.ret, percentage = FALSE))
  wpmd<-  Weights(PMD(qi.ret, percentage = FALSE))
  wpmtd<- Weights(PMTD(qi.ret, percentage = FALSE))
  V <- cov(qi.ret)
  werc<-Weights(PERC(V, percentage = FALSE))
  werc.Bl<-Weights(PERC(BLR@posteriorCovar, percentage = FALSE))
  #PPrior <- tangencyPortfolio(data = DataSub, spec = MSPrior,
  #                            constraints = BoxC)
  #PBl <- tangencyPortfolio(data = DataSub, spec = MSBl,
  #                         constraints = BoxC)
  Weights <- rbind(getWeights(PPrior), getWeights(PBl), Weights(cd), wgmv, wpmd, wpmtd, werc, werc.Bl)
  #colnames(Weights) <- ANames
  rownames(Weights) <- c("Prior", "BL", "cd", "GMV", "MD", "MTD", "ERC", "ERCBL")
  return(Weights)
}  

## Conducting back test
Backtest <- list()
i=1
#i=156
cum.n=180
for(i in 1:length(idx)){
  #DataSub <- window(R, start = start(AssetsM), end = idx[i])
  from<-sample.mon[i]
  to<-sample.mon[i+11]
  insample.range<-paste(from, to, sep="/")
  #outsample.m = "2004-01", the first list element of the data
  outsample.m<-sample.mon[i+12]
  yr.qi.ts<-tw50.hist.prices[[outsample.m]]
  yr.qi.ts<-yr.qi.ts[paste("/", to, sep="")]
  #yr.qi.ts.1<-coredata(yr.qi.ts)
  # compute daily returns
  #qi.ret<-(yr.qi.ts - lag(yr.qi.ts, 1))/yr.qi.ts 
  #qi.ret<-qi.ret[-1,]
  # use 250 daily returns to compute covariance matrix
  #DataSub<-tail(qi.ret, 250)
  DataSub<-yr.qi.ts
  BLR <- PostDist[[i]]
  Backtest[[i]] <- BT(DataSub, BLR, cum.n=180)
}
Backtest[[1]]
rowSums(Backtest[[1]])
#===============================================================================
# arrange monthly returns for component stocks
#================================================================================
insample.range<-as.Date(seq(as.yearmon("2004-01-31"),as.yearmon("2016-12-31"),
                            by=1/12), frac=1)
insample.range<-format(insample.range, "%Y-%m")
a<-tw50.hist.prices[[1]]
b<-a[endpoints(a, on="months", k=1), ]
b.ret<-returnSeries(b)
b.ret<-as.xts(b.ret)
outsam.ret<-matrix(NA, nrow = length(insample.range), ncol=50)
i="2004-01"
for (i in 1:length(insample.range)){
  jj<-insample.range[i]
  a<-tw50.hist.prices[[jj]]
  b<-a[endpoints(a, on="months", k=1), ]
  b.ret<-returnSeries(b)
  b.ret<-as.xts(b.ret)
  outsam.ret[i,]<-b.ret[jj]
}
outsam.ts<-timeSeries(outsam.ret, index = time(WPrior))
#=========================================================================
# compute benchmark returns
#=========================================================================
etf50.ts<-xts(as.numeric(etf50.df[,2]), order.by = as.Date(etf50.df[,1], format="%Y%m%d"))
etf50.ts<-etf50.ts["2003-12/2016-12"]
etf50.m<-etf50.ts[endpoints(etf50.ts, on="months", k=1), ]
etf.ret<-returnSeries(etf50.m)
etf.ret<-etf.ret[-1,]
etf.ret<-timeSeries(etf.ret, charvec = names(etf.ret))
Portetf<-cumprod(etf.ret+1)
#==============================================================================
## Extracting weights based on Prior mean (market equilibrium) and covariance  
#===============================================================================
insample.range<-as.Date(seq(as.yearmon("2004-01-31"),as.yearmon("2016-12-31"),
                            by=1/12), frac=1)
insample.range1<-as.Date(seq(as.yearmon("2003-12-31"),as.yearmon("2016-12-31"),
                             by=1/12), frac=1)
weight.name<-c("Prior", "BL", "cd", "GMV", "MD", "MTD", "ERC", "ERCBL")
portmkv<-list()
portret<-list()
# performance analysis
var<-list()
es<-list() # expected shortfall
sr<-list() # sharpe ratio
ra<-list() # annualized returns
dd<-list() # drawdowns


i=1
for (i in 1:length(weight.name)){
     weight.i <- matrix(unlist(lapply(Backtest, function(x) x[weight.name[i], ])), 
                        ncol = NAssets, byrow = TRUE)
     #w.i<-paste("W", i, sep = "")
     #port.i<-paste("Port", i, sep="")
     weight.i<-zoo(weight.i, order.by = sample.mon[13:168]) 
     # convert string into variable name: eval(parse(text=w.i))
     RetFac <- 1 + rowSums(weight.i * outsam.ts) 
     # RetFac[1] <- 100
     RetFac<-c(100,RetFac) 
     #insample.range<-c("2003-12", insample.range)
     # PortPrior
     portmkv[[i]]<- timeSeries(cumprod(RetFac), charvec = insample.range1)
     # RetPortPrior
     portret[[i]]<-portmkv[[i]]/lag(portmkv[[i]],1)-1
     portret[[i]]<-portret[[i]][-1]
     portret[[i]]<-xts(portret[[i]], order.by = insample.range)
     #
     var[[i]]<- -100 * VaR(portret[[i]], p = 0.95, method = "gaussian")
     es[[i]]<- -100 * ES(portret[[i]], p = 0.95, method = "gaussian")
     sr[[i]]<- SharpeRatio(portret[[i]])
     ra[[i]]<- Return.annualized(portret[[i]], scale = 12)
     dd[[i]]<- -100 * findDrawdowns(portret[[i]])$return
}
var[[length(weight.name)+1]]<- -100 * VaR(etf.ret, p = 0.95, method = "gaussian")
es[[length(weight.name)+1]]<- -100 * ES(etf.ret, p = 0.95, method = "gaussian")
sr[[length(weight.name)+1]]<- SharpeRatio(etf.ret)
ra[[length(weight.name)+1]]<- Return.annualized(etf.ret, scale = 12)
dd[[length(weight.name)+1]]<- -100 * findDrawdowns(etf.ret)$return
# drawdowns summary (Ref: Pfaff, Table 12.3, p252)
dd.summary<-list()
dd.times<-list()
i=1
for (i in 1:(length(weight.name)+1)){
  dd.summary[[i]]<-dd[[i]][dd[[i]]!=0]
  dd.times[[i]]<-length(dd.summary[[i]])
  dd.summary[[i]]<-c(no = dd.times[[i]],summary(dd.summary[[i]]))
}

summary.part1<-matrix(unlist(dd.summary), nrow=7)
rownames(summary.part1)<-names(dd.summary[[1]])
colnames(summary.part1)<-c(weight.name, "ETF")
summary.part1
print(xtable(summary.part1), include.rownames = TRUE, include.colnames = TRUE, 
      sanitize.text.function = I)
#
summary.part2<-matrix(rep(0, 36), nrow=4)
# i=1
# for (i in 1:4){
#     nm<-c("var", "es", "sr", "ra")
#     summary.part2[i,]<-unlist(eval(parse(text=nm[i])))
# }
summary.part2[1,]<-unlist(var)
summary.part2[2,]<-unlist(es)
temp.sr<-matrix(unlist(sr), ncol=3, byrow = TRUE)
summary.part2[3,]<-temp.sr[,1]
summary.part2[4,]<-unlist(ra)
colnames(summary.part2)<-c(weight.name, "ETF")
rownames(summary.part2)<-c("VaR", "ES", "SR", "RA")
summary.part2
class(summary.part2)
t(summary.part2)
#names(portret)<-c(weight.name, "ETF50")
print(xtable(summary.part2), include.rownames = TRUE, include.colnames = TRUE, 
             sanitize.text.function = I)


## Portfolio statistics 
## VaR 
#PriorVAR <- -100 * VaR(portret.ts, p = 0.95, method = "gaussian") 
#BLVAR <- -100 * VaR(RetBL, p = 0.95, method = "gaussian")





#===========================================================================
## Extracting weights based BL posterior
# Wbl <- matrix(unlist(lapply(Backtest, function(x) x[2, ])), 
#                  ncol = NAssets, byrow = TRUE)
# Wbl <- zoo(Wbl, order.by = sample.mon[13:168]) 
# ## Extracting weights based cd
# Wcd <- matrix(unlist(lapply(Backtest, function(x) x[3, ])), 
#               ncol = NAssets, byrow = TRUE)
# Wcd <- zoo(Wcd, order.by = sample.mon[13:168]) 
# ## Extracting weights based md (most diversified portfolio)
# Wmd <- matrix(unlist(lapply(Backtest, function(x) x[4, ])), 
#               ncol = NAssets, byrow = TRUE)
# Wmd <- zoo(Wmd, order.by = sample.mon[13:168]) 
# ## Extracting weights based mtd (minimum tail-dependent portfolio)
# Wmtd <- matrix(unlist(lapply(Backtest, function(x) x[5, ])), 
#               ncol = NAssets, byrow = TRUE)
# Wmtd <- zoo(Wmtd, order.by = sample.mon[13:168]) 
# ## Extracting weights from ERC based on prior covariance matrix 
# Werc <- matrix(unlist(lapply(Backtest, function(x) x[6, ])), 
#                ncol = NAssets, byrow = TRUE)
# Werc <- zoo(Werc, order.by = sample.mon[13:168]) 
# ## Extracting weights from ERC based on posterior covariance matrix 
# Werc.bl <- matrix(unlist(lapply(Backtest, function(x) x[7, ])), 
#                ncol = NAssets, byrow = TRUE)
# Werc.bl <- zoo(Werc.bl, order.by = sample.mon[13:168]) 
#=================================================================

#outsam.ts<-xts(outsam.ret, order.by = as.Date(insample.range, format="%Y-%m"))








#===============================================
# create weight lists to store optimized weights
# prior = use historical mean returns and covariance matrices
# post = use BL model market equilibrium returns and adjusted investors' views 
# covariance matrices
#-------------------------------------------------------
# wGMV<-list()
# wGMVbl<-list()
# wERC<-list() #equal risk contribution
# wERCbl<-list()
# wMDP<-list() # most diversified portfolio
# wMDPbl<-list()
# wMTD<-list() # minimum tail dependent portfolio
# wMTDbl<-list()

#conducting backtesting
# global minimum variance portfolio
wGMV.mtx<-matrix(NA, ncol=50, nrow=outsample.n)
# most diversified portfolio
wMD.mtx<-wGMV.mtx
# equal risk contribution portfolio
wERC.mtx<-wGMV.mtx
# minimum tail dependence portfolio
wMTD.mtx<-wGMV.mtx
#
pspec <- portfolioSpec() 
gmv <- pspec
i=76 # i=96
i=1
#library(SIT)
for (i in 1:sample.n){
     sample.range<-paste(yr.mon[insample.start.mon+i-1],"/", sep="")
     sample.range<-paste(sample.range, yr.mon[insample.end.mon+i-1], sep="")
     #day.ret<-diff(log(tw50.hist.prices[[i]]))
     temp.prices<-tw50.hist.prices[[i]][sample.range]
     day.ret<-returnseries(as.timeSeries(temp.prices), method="discrete", trim=TRUE)
     #==============================================================================
     #below we try to use systematic investor tool box
     #==============================================================================
     ia = create.historical.ia(day.ret, 252)
     s0 = apply(coredata(day.ret),2,sd)     
     ia$correlation = cor(coredata(day.ret), use='complete.obs',method='pearson')
     ia$cov = ia$correlation * (s0 %*% t(s0))
     
     
     # here ends the systematic investor tool codes
     #==============================================================================
     day.ret<-as.timeSeries(day.ret)
     hist.mean.i<-colMeans(day.ret, na.rm=TRUE)
     V<-cov(day.ret)
     #V<-factor.model.stat(day.ret)
     #V<-cov(day.ret)
     # Weights(PGMV(day.ret))
     #gmvpf<-minvariancePortfolio(data = day.ret, spec = gmv, constraints = "LongOnly")
     #wGMV.mtx[i, ] <- c(getWeights(gmvpf))
     wGMV.mtx[i, ] <- Weights(PGMV(day.ret))
     wMD.mtx[i,]   <-Weights(PMD(day.ret))
     wERC.mtx[i,]  <-Weights(PERC(V))
     wMTD.mtx[i,]  <-Weights(PMTD(day.ret))
}




#先將週資料檔案之名稱讀入; 
file_path="D:/Data TWN50/日期排列(週).txt"
file_ind=read.table(file=file_path)
num_rows=dim(file_ind)[1]
num_rows # number of weeks in the sample data
# a=file_ind[2,1]

#先將月資料檔案之名稱讀入; 
file_pathm="D:/Data TWN50/日期排列(月).txt"
file_indm=read.table(file=file_pathm)
num_rowsm=dim(file_indm)[1]
num_rowsm

setwd("D:/亞洲大學碩士班指導論文/筌鈞/R script")

path="D:/Data TWN50/新表格(週)/"
pathm="D:/Data TWN50/新表格(月)/"
# Weekly basis: cov1_100 means using reterns with percentage to compute the covaraince matrix;
cov1_path=paste(path, "cov1_", sep="")
cov2_path=paste(path, "cov2_", sep="")
mean_path=paste(path, "mean_", sep="")

twn50_ret=paste(path, "twn50_ret.txt", sep="")
twn50_id =paste(path, "twn50_id.txt", sep="")

pick_5ret =paste(path, "pick_5ret", sep="")
pick_10ret =paste(path, "pick_10ret", sep="")
pick_15ret =paste(path, "pick_15ret", sep="")
pick_20ret =paste(path, "pick_20ret", sep="")
pick_25ret =paste(path, "pick_25ret", sep="")
pick_50ret =paste(path, "pick_50ret", sep="")
pick_100ret =paste(path, "pick_100ret", sep="")
pick_150ret =paste(path, "pick_150ret", sep="")
pick_200ret =paste(path, "pick_200ret", sep="")
pick_allret =paste(path, "pick_allret", sep="")

# Monthly basis:
cov1_pathm=paste(pathm, "cov1_", sep="")
cov2_pathm=paste(pathm, "cov2_", sep="")
mean_pathm=paste(pathm, "mean_", sep="")

twn50m_ret=paste(pathm, "twn50_ret.txt", sep="")
twn50m_id =paste(pathm, "twn50_id.txt", sep="")

pick_1mret =paste(pathm, "pick_1mret", sep="")
pick_3mret =paste(pathm, "pick_3mret", sep="")
pick_6mret =paste(pathm, "pick_6mret", sep="")
pick_9mret =paste(pathm, "pick_9mret", sep="")
pick_12mret =paste(pathm, "pick_12mret", sep="")
pick_24mret =paste(pathm, "pick_24mret", sep="")
pick_36mret =paste(pathm, "pick_36mret", sep="")
pick_48mret =paste(pathm, "pick_48mret", sep="")
pick_allmret =paste(pathm, "pick_allmret", sep="")

#out2_mean=paste(path, "mean2.txt", sep="")
#i=1
#i=2

# twn50.ls 為每週台灣五十之股票代號; 
twn50.ls<-list()
# twn50.ret.ls 為每週台灣五十之股票報酬率;
twn50.ret.ls<-list()

# twn50m.ls 為每月台灣五十之股票代號;
twn50m.ls<-list()
# twn50m.ret.ls 為每月台灣五十之股票報酬率;
twn50m.ret.ls<-list()

##########################################################################
#以下為週報酬率之設定;
##########################################################################
#pick.5ret.ls 為台灣五十股票最近5週累計報酬率所形成的pick matrix
pick.5ret.ls<-list()
#pick.10ret.ls 為台灣五十股票最近10週累計報酬率所形成的pick matrix
pick.10ret.ls<-list()
#pick.15ret.ls 為台灣五十股票最近15週累計報酬率所形成的pick matrix
pick.15ret.ls<-list()
#pick.20ret.ls 為台灣五十股票最近20週累計報酬率所形成的pick matrix
pick.20ret.ls<-list()
#pick.25ret.ls 為台灣五十股票最近25週累計報酬率所形成的pick matrix
pick.25ret.ls<-list()
#pick.50ret.ls 為台灣五十股票最近50週累計報酬率所形成的pick matrix
pick.50ret.ls<-list()
#pick.100ret.ls 為台灣五十股票最近100週累計報酬率所形成的pick matrix
pick.100ret.ls<-list()
#pick.150ret.ls 為台灣五十股票最近150週累計報酬率所形成的pick matrix
pick.150ret.ls<-list()
#pick.200ret.ls 為台灣五十股票最近200週累計報酬率所形成的pick matrix
pick.200ret.ls<-list()
#pick.ret.ls 為台灣五十股票最近5年累計報酬率所形成的pick matrix
pick.ret.ls<-list()

##########################################################################
#以下為月報酬率之設定;
##########################################################################
#pick.1mret.ls 為台灣五十股票最近1月累計報酬率所形成的pick matrix
pick.1mret.ls<-list()
#pick.3mret.ls 為台灣五十股票最近3月累計報酬率所形成的pick matrix
pick.3mret.ls<-list()
#pick.6mret.ls 為台灣五十股票最近6月累計報酬率所形成的pick matrix
pick.6mret.ls<-list()
#pick.9mret.ls 為台灣五十股票最近9月累計報酬率所形成的pick matrix
pick.9mret.ls<-list()
#pick.12mret.ls 為台灣五十股票最近12月累計報酬率所形成的pick matrix
pick.12mret.ls<-list()
#pick.24mret.ls 為台灣五十股票最近24月累計報酬率所形成的pick matrix
pick.24mret.ls<-list()
#pick.36mret.ls 為台灣五十股票最近36月累計報酬率所形成的pick matrix
pick.36mret.ls<-list()
#pick.48mret.ls 為台灣五十股票最近48月累計報酬率所形成的pick matrix
pick.48mret.ls<-list()
#pick.retm.ls 為台灣五十股票最近5年累計報酬率所形成的pick matrix
pick.retm.ls<-list()

########################################################################
# 以下計算weekly basis 之平均數及共變異數矩陣 
########################################################################
######################################################################
# create Black Litterman's views based on momentum strategy
# Momentum strategy is formed by using the past accumulated returns: 
# 5wks, 10wks, 15wks, 20wks, 25wks, 50wks, 100wks, 150wks, 200wks, 
# all historical periods.
######################################################################
# qq 代表選取股票累積報酬率前(1-qq)百分比的股票;
# qq=0.8，代表前20%的股票;
#############################################################

i=1
j=1
k=1
qq=c(0.7,0.8,0.9)

for (k in 1:3)  { # k 代表 qq
  
  for (i in 1:num_rows) {
    file=paste(file_ind[i,1], ".xlsx",sep="")
    file1=paste(path, file, sep="")
    wb <- loadWorkbook(file1,create = FALSE)
    data<- readWorksheet(wb, sheet="Sheet1", header=FALSE)
    dataid<-data[1,-1] #  刪除第一欄，因為該資料為"年月日" 
    stockid<- apply(dataid, 1, FUN=function(x) substr(x,11,14))
    twn50.ls[[i]]<-stockid
    name.ls.id<- data[dim(data)[1],1] 
    names(twn50.ls)[i]<- name.ls.id
    
    #將多餘文字在第一列刪除;
    data<-data[-1,]
    #將第一行日期刪除;
    data1<-data[,-1]
    tmp.mtx<-matrix(as.numeric(unlist(data1)), ncol=50)
    tmp.ret<-diff(log(tmp.mtx))*100
    
    if (k==1) {  #此條件主要讓迴圈j只重覆做一個1:num_rows次，不是k次;
      for (j in 1:4) {  #計算四種期間的歷史變異數及平均數: 100wks, 150wks, 200wks, 250wks;
        # 將文字資料轉為數字
        temp <-lapply(tail(data1,50*(j+1)), as.numeric)
        temp.df <- do.call("cbind", temp)
        temp.ret<-diff(log(temp.df))*100
        # 將資料轉為time series; 
        temp.ts<-timeSeries(temp.df, as.Date(tail(data[,1],50*(j+1))))
        data.ret<-diff(log(temp.ts))*100
        # class(data.ret)
        # 利用 BurStFin套件計算 covariance matrix and mean vector;
        data.fmvar <- factor.model.stat(temp.ret)
        data.fmvar1<- var.shrink.eqcor(temp.ret)
        mean1<- matrix(colMeans(temp.ret, na.rm=TRUE),nrow=1)
        cov1_pathj=paste(paste(cov1_path, 50*(j+1), sep=""), ".txt", sep="")
        cov2_pathj=paste(paste(cov2_path, 50*(j+1), sep=""), ".txt", sep="")
        mean_pathj=paste(paste(mean_path, 50*(j+1), sep=""), ".txt", sep="")
        
        write.table(mean1, file=mean_pathj, append=TRUE, row.names=FALSE, col.names=FALSE)
        write.table(data.fmvar, file=cov1_pathj, append=TRUE, row.names = FALSE, col.names=FALSE)
        write.table(data.fmvar1, file=cov2_pathj, append=TRUE, row.names = FALSE, col.names=FALSE)
        rm (data.fmvar, data.fmvar1, mean1)
      }
    }
    
    
    
    #計算累計報酬率並取出最近N期累計報酬率; N=200代表最近過去200wks報酬率;
    accum.all.ret<-colSums(tmp.ret,na.rm=TRUE)
    accum.200.ret<-colSums(tail(tmp.ret,200),na.rm=TRUE)
    accum.150.ret<-colSums(tail(tmp.ret,150),na.rm=TRUE)
    accum.100.ret<-colSums(tail(tmp.ret,100), na.rm=TRUE)
    accum.50.ret<-colSums(tail(tmp.ret,50), na.rm=TRUE)
    accum.25.ret<-colSums(tail(tmp.ret,25), na.rm=TRUE)
    accum.20.ret<-colSums(tail(tmp.ret,20), na.rm=TRUE)
    accum.15.ret<-colSums(tail(tmp.ret,15), na.rm=TRUE)
    accum.10.ret<-colSums(tail(tmp.ret,10), na.rm=TRUE)
    accum.5.ret<-colSums(tail(tmp.ret,5), na.rm=TRUE)
    
    last.day.ret<-tail(tmp.ret,1)
    twn50.ret.ls[[i]]<-as.vector(last.day.ret)
    names(twn50.ret.ls)[i]<-name.ls.id
    ######################################################################
    # create Black Litterman's views based on momentum strategy
    # Momentum strategy is formed by using the past accumulated returns: 
    # 5wks, 10wks, 15wks, 20wks, 25wks, 50wks, 100wks, 150wks, 200wks, 
    # all historical periods.
    ######################################################################
    # q 代表選取股票累積報酬率前(1-q)百分比的股票;
    # q=0.8，代表前20%的股票;
    #############################################################
    # 1. create Pick matrix for 5 trading weeks;
    #############################################################
    q=qq[k]
    q.5<-quantile(accum.5.ret,q)
    qq.5<-accum.5.ret>=q.5
    qq.5.mtx<-matrix(as.numeric(qq.5),ncol=50)
    # nw: number of winners (q=80% quantile)
    nw<-sum(qq.5)
    pick.5<- apply(qq.5.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.5ret.ls[[i]]<-as.vector(pick.5)
    
    #############################################################
    # 2. create Pick matrix for 10 trading weeks;
    #############################################################
    q.10<-quantile(accum.10.ret,q)
    qq.10<-accum.10.ret>=q.10
    qq.10.mtx<-matrix(as.numeric(qq.10),ncol=50)
    # nw: number of winners (q=80% quantile)
    nw<-sum(qq.10)
    pick.10<- apply(qq.10.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.10ret.ls[[i]]<-as.vector(pick.10)
    
    #############################################################
    # 3. create Pick matrix for 15 trading weeks;
    #############################################################
    q.15<-quantile(accum.15.ret,q)
    qq.15<-accum.15.ret>=q.15
    qq.15.mtx<-matrix(as.numeric(qq.15),ncol=50)
    # nw: number of winners (q=80% quantile)
    nw<-sum(qq.15)
    pick.15<- apply(qq.15.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.15ret.ls[[i]]<-as.vector(pick.15)
    
    #############################################################
    # 4. create Pick matrix for 20 trading weeks;
    #############################################################
    q.20<-quantile(accum.20.ret,q)
    qq.20<-accum.20.ret>=q.20
    qq.20.mtx<-matrix(as.numeric(qq.20),ncol=50)
    # nw: number of winners (q=80% quantile)
    nw<-sum(qq.20)
    pick.20<- apply(qq.20.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.20ret.ls[[i]]<-as.vector(pick.20)
    
    
    #############################################################
    # 5. create Pick matrix for 0.5 year or 25 trading weeks;
    #############################################################
    #q=0.8
    q.25<-quantile(accum.25.ret,q)
    qq.25<-accum.25.ret>=q.25
    qq.25.mtx<-matrix(as.numeric(qq.25),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.25)
    pick.25<- apply(qq.25.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.25ret.ls[[i]]<-as.vector(pick.25)
    
    #############################################################
    # 6. create Pick matrix for 1 year or 50 trading weeks;
    #############################################################
    #q=0.8
    q.50<-quantile(accum.50.ret,q)
    qq.50<-accum.50.ret>=q.50
    qq.50.mtx<-matrix(as.numeric(qq.50),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.50)
    pick.50<- apply(qq.50.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.50ret.ls[[i]]<-as.vector(pick.50)
    
    #############################################################
    # 7. create Pick matrix for 2 year or 100 trading weeks;
    #############################################################
    #q=0.8
    q.100<-quantile(accum.100.ret,q)
    qq.100<-accum.100.ret>=q.100
    qq.100.mtx<-matrix(as.numeric(qq.100),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.100)
    pick.100<- apply(qq.100.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.100ret.ls[[i]]<-as.vector(pick.100)
    #############################################################
    # 8. create Pick matrix for 3 year or 150 trading weeks;
    #############################################################
    #q=0.8
    q.150<-quantile(accum.150.ret,q)
    qq.150<-accum.150.ret>=q.150
    qq.150.mtx<-matrix(as.numeric(qq.150),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.150)
    pick.150<- apply(qq.150.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.150ret.ls[[i]]<-as.vector(pick.150)
    
    #############################################################
    # 9. create Pick matrix for 4 year or 200 trading weeks;
    #############################################################
    q.200<-quantile(accum.200.ret,q)
    qq.200<-accum.200.ret>=q.200
    qq.200.mtx<-matrix(as.numeric(qq.200),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.200)
    pick.200<- apply(qq.200.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.200ret.ls[[i]]<-as.vector(pick.200)
    
    #############################################################
    #10. create Pick matrix for 5-year trading days;
    #############################################################   
    q.all<-quantile(accum.all.ret,q)
    qq.all<-accum.all.ret>=q.all
    qq.all.mtx<-matrix(as.numeric(qq.all),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.all)
    pick.all<- apply(qq.all.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.ret.ls[[i]]<-as.vector(pick.all)
    ###################################################################
    
  }
  
  #將pick資料合併為matrix資料型態;50檔股票每週投資權重
  pick.5ret.df<- do.call("rbind", pick.5ret.ls)
  pick.10ret.df<- do.call("rbind", pick.10ret.ls)
  pick.15ret.df<- do.call("rbind", pick.15ret.ls)
  pick.20ret.df<- do.call("rbind", pick.20ret.ls)
  pick.25ret.df<- do.call("rbind", pick.25ret.ls) 
  pick.50ret.df<- do.call("rbind", pick.50ret.ls)
  pick.100ret.df<- do.call("rbind", pick.100ret.ls) 
  pick.150ret.df<- do.call("rbind", pick.150ret.ls)   
  pick.200ret.df<- do.call("rbind", pick.200ret.ls)
  pick.allret.df<- do.call("rbind", pick.ret.ls)   
  
  q=q*100
  
  pick_5retq=paste(paste(pick_5ret, q, sep=""), ".txt", sep="")
  pick_10retq=paste(paste(pick_10ret, q, sep=""), ".txt", sep="")
  pick_15retq=paste(paste(pick_15ret, q, sep=""), ".txt", sep="")
  pick_20retq=paste(paste(pick_20ret, q, sep=""), ".txt", sep="")
  pick_25retq=paste(paste(pick_25ret, q, sep=""), ".txt", sep="")
  pick_50retq=paste(paste(pick_50ret, q, sep=""), ".txt", sep="")
  pick_100retq=paste(paste(pick_100ret, q, sep=""), ".txt", sep="")
  pick_150retq=paste(paste(pick_150ret, q, sep=""), ".txt", sep="")
  pick_200retq=paste(paste(pick_200ret, q, sep=""), ".txt", sep="")
  pick_allretq=paste(paste(pick_allret, q, sep=""), ".txt", sep="")
  
  write.table(pick.5ret.df, file=pick_5retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.10ret.df, file=pick_10retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.15ret.df, file=pick_15retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.20ret.df, file=pick_20retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.25ret.df, file=pick_25retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.50ret.df, file=pick_50retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.100ret.df, file=pick_100retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.150ret.df, file=pick_150retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.200ret.df, file=pick_200retq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.allret.df, file=pick_allretq, append=TRUE, row.names=FALSE, col.names=FALSE)
}

#將台灣五十每週報酬率之list的資料型態，轉為dataframe或matrix之資料型態;
#twn50.ret.df 為 392X50 之矩陣資料, 第一列代表在2003/07/04之台灣五十股票之週報酬率; 
twn50.ret.df <- do.call("rbind", twn50.ret.ls)
#twn50.ls[[i]]為1X50之資料，代表每週台灣五十股票代碼;
#twn50.df 為50x392之矩陣資料，第1欄代表2003/07/04之台灣五十股票代號，第2欄代表2003/07/11之台灣五十股票代號… 
twn50.df<- do.call("cbind", twn50.ls)
colnames(twn50.df)<-rownames(twn50.ret.df)
write.table(twn50.ret.df, file=twn50_ret, append=TRUE, row.names=TRUE, col.names=FALSE)
write.table(twn50.df, file=twn50_id, append=TRUE, row.names=FALSE, col.names=TRUE)    


###############################################################################################################

########################################################################
# 以下計算monthly basis 之平均數及共變異數矩陣 
########################################################################
k=1
i=1
j=1

qm=c(0.7,0.8,0.9)

for (k in 1:3)  { # k 代表 qm
  
  for (i in 1:num_rowsm) {
    file=paste(file_indm[i,1], ".xlsx",sep="")
    file1=paste(pathm, file, sep="")
    wb <- loadWorkbook(file1,create = FALSE)
    data<- readWorksheet(wb, sheet="Sheet1", header=FALSE)
    dataid<-data[1,-1] #  刪除第一欄，因為該資料為"年月日" 
    stockid<- apply(dataid, 1, FUN=function(x) substr(x,11,14))
    twn50m.ls[[i]]<-stockid
    name.ls.id<- data[dim(data)[1],1] 
    names(twn50m.ls)[i]<- name.ls.id
    
    #將多餘文字在第一列刪除;
    data<-data[-1,]
    #將第一行日期刪除;
    data1<-data[,-1]
    #purpose: 將最後一筆報酬率取出; 
    #tmp<-tail(data1)
    tmp.mtx<-matrix(as.numeric(unlist(data1)), ncol=50)
    tmp.ret<-diff(log(tmp.mtx))*100
    #計算累計報酬率; 6 代表最近過去6 months 報酬率;
    
    if (k==1) {  #此條件主要讓迴圈j只重覆做1:num_rows次
      for (j in 1:4) {  #計算四種期間的歷史變異數及平均數: 12m, 24m, 36m, 48m;
        # 將文字資料轉為數字
        temp <-lapply(tail(data1,12*j), as.numeric)
        temp.df <- do.call("cbind", temp)
        temp.ret<-diff(log(temp.df))*100
        # 將資料轉為time series; 
        # temp.ts<-timeSeries(temp.df, as.Date(tail(data[,1],12*j)))
        # temp.ts<-timeSeries(tail(temp.df, 12*j))
        temp.test<-tail(temp.df, 12*j)
        test.ret<-diff(log(temp.test))*100
        # class(data.ret)
        # 利用 BurStFin套件計算 covariance matrix and mean vector;
        data.fmvar <- factor.model.stat(test.ret)
        data.fmvar1<- var.shrink.eqcor(test.ret)
        mean1<- matrix(colMeans(test.ret, na.rm=TRUE),nrow=1)
        cov1_pathmj=paste(paste(cov1_pathm, 12*j, sep=""), ".txt", sep="")
        cov2_pathmj=paste(paste(cov2_pathm, 12*j, sep=""), ".txt", sep="")
        mean_pathmj=paste(paste(mean_pathm, 12*j, sep=""), ".txt", sep="")
        
        write.table(mean1, file=mean_pathmj, append=TRUE, row.names=FALSE, col.names=FALSE)
        write.table(data.fmvar, file=cov1_pathmj, append=TRUE, row.names = FALSE, col.names=FALSE)
        write.table(data.fmvar1, file=cov2_pathmj, append=TRUE, row.names = FALSE, col.names=FALSE)
        rm (data.fmvar, data.fmvar1, mean1)
      }
    }
    
    accum.allm.ret<-colSums(tmp.ret,na.rm=TRUE)
    accum.48m.ret<-colSums(tail(tmp.ret,48),na.rm=TRUE)
    accum.36m.ret<-colSums(tail(tmp.ret,36),na.rm=TRUE)
    accum.24m.ret<-colSums(tail(tmp.ret,24), na.rm=TRUE)
    accum.12m.ret<-colSums(tail(tmp.ret,12), na.rm=TRUE)
    accum.9m.ret<-colSums(tail(tmp.ret,9), na.rm=TRUE)
    accum.6m.ret<-colSums(tail(tmp.ret,6), na.rm=TRUE)
    accum.3m.ret<-colSums(tail(tmp.ret,3), na.rm=TRUE)
    accum.1m.ret<-colSums(tail(tmp.ret,1), na.rm=TRUE)
    
    last.day.ret<-tail(tmp.ret,1)
    last.day.ret
    twn50m.ret.ls[[i]]<-as.vector(last.day.ret)
    names(twn50m.ret.ls)[i]<-name.ls.id
    
    
    #############################################################
    # create Black Litterman's views based on momentum strategy
    # Momentum strategy is formed by using the past accumulated returns: 
    # 1, 3, 6 , 9 months, 12m, 24m, 36m, 48m, all historical periods.
    #############################################################
    # 1. create Pick matrix for 1 trading month;
    #############################################################
    q=qm[k]
    q.1m<-quantile(accum.1m.ret,q)
    qq.1m<-accum.1m.ret>=q.1m
    qq.1m.mtx<-matrix(as.numeric(qq.1m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.1m)
    pick.1m<- apply(qq.1m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.1mret.ls[[i]]<-as.vector(pick.1m)
    #############################################################
    # 2. create Pick matrix for 3 trading months;
    #############################################################
    
    q.3m<-quantile(accum.3m.ret,q)
    qq.3m<-accum.3m.ret>=q.3m
    qq.3m.mtx<-matrix(as.numeric(qq.3m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.3m)
    pick.3m<- apply(qq.3m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.3mret.ls[[i]]<-as.vector(pick.3m)
    
    #############################################################
    # 3. create Pick matrix for 6 trading months;
    #############################################################
    
    q.6m<-quantile(accum.6m.ret,q)
    qq.6m<-accum.6m.ret>=q.6m
    qq.6m.mtx<-matrix(as.numeric(qq.6m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.6m)
    pick.6m<- apply(qq.6m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.6mret.ls[[i]]<-as.vector(pick.6m)
    
    #############################################################
    # 4. create Pick matrix for 9 trading months;
    #############################################################
    
    q.9m<-quantile(accum.9m.ret,q)
    qq.9m<-accum.9m.ret>=q.9m
    qq.9m.mtx<-matrix(as.numeric(qq.9m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.9m)
    pick.9m<- apply(qq.9m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.9mret.ls[[i]]<-as.vector(pick.9m)
    
    #############################################################
    # 5. create Pick matrix for 12 trading months;
    #############################################################
    
    q.12m<-quantile(accum.12m.ret,q)
    qq.12m<-accum.12m.ret>=q.12m
    qq.12m.mtx<-matrix(as.numeric(qq.12m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.12m)
    pick.12m<- apply(qq.12m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.12mret.ls[[i]]<-as.vector(pick.12m)
    
    #############################################################
    # 6. create Pick matrix for 24 trading months;
    #############################################################
    #q=0.8
    q.24m<-quantile(accum.24m.ret,q)
    qq.24m<-accum.24m.ret>=q.24m
    qq.24m.mtx<-matrix(as.numeric(qq.24m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.24m)
    pick.24m<- apply(qq.24m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.24mret.ls[[i]]<-as.vector(pick.24m)
    
    #############################################################
    # 7. create Pick matrix for 36 trading months;
    #############################################################
    #q=0.8
    q.36m<-quantile(accum.36m.ret,q)
    qq.36m<-accum.36m.ret>=q.36m
    qq.36m.mtx<-matrix(as.numeric(qq.36m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.36m)
    pick.36m<- apply(qq.36m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.36mret.ls[[i]]<-as.vector(pick.36m)
    
    #############################################################
    # 8. create Pick matrix for 48 trading months;
    #############################################################
    q.48m<-quantile(accum.48m.ret,q)
    qq.48m<-accum.48m.ret>=q.48m
    qq.48m.mtx<-matrix(as.numeric(qq.48m),ncol=50)
    # nw: number of winners (80% quantile)
    nw<-sum(qq.48m)
    pick.48m<- apply(qq.48m.mtx, 2, FUN=function(x) if (x==1) x=1/nw else x=1/(50-nw))
    pick.48mret.ls[[i]]<-as.vector(pick.48m)
    
    #############################################################
    # 9. create Pick matrix for 5-year trading months;
    #############################################################   
    q.allm<-quantile(accum.allm.ret,q)
    qq.allm<-accum.allm.ret>=q.allm
    qq.allm.mtx<-matrix(as.numeric(qq.allm),ncol=50)
    # nw: number of winners (80% quantile)
    nwm<-sum(qq.allm)
    pick.allm<- apply(qq.allm.mtx, 2, FUN=function(x) if (x==1) x=1/nwm else x=1/(50-nwm))
    pick.retm.ls[[i]]<-as.vector(pick.allm)
    ###################################################################
  }   
  
  #將pick資料合併為matrix資料型態;
  pick.1mret.df<- do.call("rbind", pick.6mret.ls) 
  pick.3mret.df<- do.call("rbind", pick.6mret.ls) 
  pick.6mret.df<- do.call("rbind", pick.6mret.ls) 
  pick.9mret.df<- do.call("rbind", pick.6mret.ls) 
  pick.12mret.df<- do.call("rbind", pick.12mret.ls)
  pick.24mret.df<- do.call("rbind", pick.24mret.ls) 
  pick.36mret.df<- do.call("rbind", pick.36mret.ls)   
  pick.48mret.df<- do.call("rbind", pick.48mret.ls)
  pick.allmret.df<- do.call("rbind", pick.retm.ls)   
  
  q=q*100
  
  pick_1mretq=paste(paste(pick_1mret, q, sep=""), ".txt", sep="")
  pick_3mretq=paste(paste(pick_3mret, q, sep=""), ".txt", sep="")
  pick_6mretq=paste(paste(pick_6mret, q, sep=""), ".txt", sep="")
  pick_9mretq=paste(paste(pick_9mret, q, sep=""), ".txt", sep="")
  pick_12mretq=paste(paste(pick_12mret, q, sep=""), ".txt", sep="")
  pick_24mretq=paste(paste(pick_24mret, q, sep=""), ".txt", sep="")
  pick_36mretq=paste(paste(pick_36mret, q, sep=""), ".txt", sep="")
  pick_48mretq=paste(paste(pick_48mret, q, sep=""), ".txt", sep="")
  pick_allmretq=paste(paste(pick_allmret, q, sep=""), ".txt", sep="")
  
  write.table(pick.1mret.df, file=pick_1mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.3mret.df, file=pick_3mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.6mret.df, file=pick_6mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.9mret.df, file=pick_9mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.12mret.df, file=pick_12mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.24mret.df, file=pick_24mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.36mret.df, file=pick_36mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.48mret.df, file=pick_48mretq, append=TRUE, row.names=FALSE, col.names=FALSE)
  write.table(pick.allmret.df, file=pick_allmretq, append=TRUE, row.names=FALSE, col.names=FALSE)
}

#將台灣五十每月報酬率之list的資料型態，轉為dataframe或matrix之資料型態;
#twn50m.ret.df 為 90X50 之矩陣資料, 第一列代表在2003/07/31之台灣五十股票之月報酬率; 
twn50m.ret.df <- do.call("rbind", twn50m.ret.ls)
# twn50.df 為50x392之矩陣資料，第1欄代表2003/07/04之台灣五十股票代號，第2欄代表2003/07/11之台灣五十股票代號… 
twn50m.df<- do.call("cbind", twn50m.ls)
colnames(twn50m.df)<-rownames(twn50m.ret.df)
write.table(twn50m.ret.df, file=twn50m_ret, append=TRUE, row.names=TRUE, col.names=FALSE)
write.table(twn50m.df, file=twn50m_id, append=TRUE, row.names=FALSE, col.names=TRUE)


###############################################################################################################










