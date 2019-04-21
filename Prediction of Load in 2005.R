library("readxl")
library(rvest)
library(dplyr)
library(forecast)
library(ggplot2)
library("openxlsx")

# get Load_KW data from excel file
Load_kW <- read_excel("Competition data.xlsx", sheet = "Load_kW")%>% filter(!is.na(Load_kW))

# create a time series object # year 8760, week 168,day 24
load_kw_y=msts(Load_kW$Load_kW,seasonal.periods = c(24,24*7,24*30,24*365),ts.frequency=8760,
             start = as.POSIXct("2002-01-01 01:00"))

# set the last 6 months as test, set previous months as traning 6*30*24=4320 points as test
#testing set = last 4320 points
ntest=6*30*24
ntrain=length(load_kw_y) - ntest

train.end=time(load_kw_y)[ntrain]
test.start=time(load_kw_y)[ntrain+1]

load_kw_y.train=window(load_kw_y,end=train.end)
load_kw_y.test=window(load_kw_y,start=test.start)



######################## Build a seasonal naive model ########################

naive_M1=snaive(load_kw_y.train,h=length(load_kw_y.test),level=95)

#autoplot(naive_M1) + autolayer(naive_M1$fitted,series="Fitted\nvalues")+
 # autolayer(load_kw_y.test,series="Testing\nset")   

accu=data.frame(accuracy(na.omit(naive_M1),load_kw_y.test))

# MAPE Training 17.73 Test 16.38


# create mape dataframe
mape_result=data.frame("naive_model"=accu$MAPE,row.names = c("test","train"))
mape_result


######################## Smoothing model ######################## Exponential smoothing state space model

#M3=stlf(load_kw_y.train, h = ntest,lambda="auto",level=95,s.window = 30)#, 
#M3F=forecast(M3,h=ntest)

#autoplot(M3$fitted,series="Fitted\nvalues")+
 # autolayer(load_kw_y.test,series="Testing\nset")#+autolayer(M3F)

#accuracy(M3F,load_kw_y.test)


# ME     RMSE       MAE         MPE      MAPE      MASE      ACF1
# Training set   -440.1741 14338.38  7954.419  -0.1075418  2.352408 0.1362683 0.6839599
# Test set     -52054.5414 93195.60 70244.017 -18.7794022 23.428768 1.2033600 0.9589556


# Residual diagnostics
#tsdisplay(M3$residuals)
#Box.test(M3$residuals,lag = 24)
#checkresiduals(M3$residuals)


######################## ARIMA model ######################## 

Acf(load_kw_y.train,lag.max = 24,main="")

ari.model=Arima(load_kw_y.train,order=c(3,1,1),lambda=0)
arima_prediction=forecast(ari.model,h=ntest)
accu=data.frame(accuracy(arima_prediction,load_kw_y.test))

# add mape result
mape_result$arima_model=accu$MAPE
mape_result
# MAPE Training 2.93 Test 27.31

#autoplot(M4F) + autolayer(M4$fitted,series="Fitted\nvalues")+
#  autolayer(load_kw_y.test,series="Testing\nset")

######################## TBATS Model ######################## Exponential Smoothing State Space Model With Box-Cox Transformation, ARMA Errors, Trend And Seasonal Components

#?tbats
tba.model=tbats(load_kw_y.train,use.box.cox=TRUE,use.trend = TRUE,
                use.arma.errors=TRUE,use.parallel=TRUE, num.cores=NULL)

tba_prediction=forecast(tba.model,h=ntest)
accu=data.frame(accuracy(tba_prediction,load_kw_y.test))

# add mape result
mape_result$tbats_model=accu$MAPE

# MAPE Training 2.60807 Test 3585.08863

######################## STL Model ######################## applying a non-seasonal forecasting method to the seasonally adjusted data and re-seasonalizing using the last year of the seasonal component.

#?stlm
stlm.model=stlm(load_kw_y.train, s.window = "periodic", robust = FALSE, method = c("arima"),
                modelfunction = NULL, model = NULL, etsmodel = "ZZN",
                lambda = NULL, biasadj = FALSE, xreg = NULL,
                allow.multiplicative.trend = FALSE)

stlm_prediction=forecast(stlm.model,h=ntest)
accu=data.frame(accuracy(stlm_prediction,load_kw_y.test))

# add mape result
mape_result$stl_model=accu$MAPE

# MAPE Training 0.986172 Test 16.324152

######################## Regression Model ######################## 
#?tslm

Load_kW_regression=data.frame(Load_kW)

Load_kW_regression$trend=1:length(Load_kW_regression$Date)
# add weekday variables for regression model
Load_kW_regression$Monday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==2,1,0)
Load_kW_regression$Tuesday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==3,1,0)
Load_kW_regression$Wednesday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==4,1,0)
Load_kW_regression$Thursday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==5,1,0)
Load_kW_regression$Friday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==6,1,0)
Load_kW_regression$Saturday=ifelse(Load_kW_regression$DayofWeek..1..Sunday..to.7..Saturday..==7,1,0)

# add month variables for regression model
Load_kW_regression$January=ifelse(Load_kW_regression$Month==1,1,0)
Load_kW_regression$February=ifelse(Load_kW_regression$Month==2,1,0)
Load_kW_regression$March=ifelse(Load_kW_regression$Month==3,1,0)
Load_kW_regression$April=ifelse(Load_kW_regression$Month==4,1,0)
Load_kW_regression$May=ifelse(Load_kW_regression$Month==5,1,0)
Load_kW_regression$June=ifelse(Load_kW_regression$Month==6,1,0)
Load_kW_regression$July=ifelse(Load_kW_regression$Month==7,1,0)
Load_kW_regression$August=ifelse(Load_kW_regression$Month==8,1,0)
Load_kW_regression$September=ifelse(Load_kW_regression$Month==9,1,0)
Load_kW_regression$October=ifelse(Load_kW_regression$Month==10,1,0)
Load_kW_regression$November=ifelse(Load_kW_regression$Month==11,1,0)




ntest=6*30*24
ntrain=length(Load_kW_regression) - ntest

reg.train=head(Load_kW_regression,ntrain)
reg.test=tail(Load_kW_regression,ntest)

reg.train.y=msts(reg.train$Load_kW,seasonal.periods = c(24,24*7,24*30,24*365),ts.frequency=8760,
                                  start = as.POSIXct("2002-01-01 01:00"))

train.Temperature=reg.train$T
train.Monday=reg.train$Monday
train.Trend=reg.train$trend
train.Hour=reg.train$Hour
train.Tuesday=reg.train$Tuesday
train.Wednesday=reg.train$Wednesday
train.Thursday=reg.train$Thursday
train.Friday=reg.train$Friday
train.Saturday=reg.train$Saturday

train.January=reg.train$January
train.February=reg.train$February
train.March=reg.train$March
train.April=reg.train$April
train.May=reg.train$May
train.June=reg.train$June
train.July=reg.train$July
train.August=reg.train$August
train.September=reg.train$September
train.October=reg.train$October
train.November=reg.train$November

train.lm.trend.season=tslm(reg.train.y~train.Temperature+train.Monday+train.Trend+train.Hour
                             +train.Tuesday+train.Wednesday+train.Thursday+train.Friday+train.Saturday+
                             train.January+train.February+train.March+train.April+train.May+
                             train.June+train.July+train.August+train.September+train.October+train.November
                             ,lambda="auto")

summary(train.lm.trend.season)

fcast = forecast(train.lm.trend.season, newdata = data.frame(
  train.Temperature=reg.test$T,
  train.Monday=reg.test$Monday,
  train.Trend=reg.test$trend,
  train.Hour=reg.test$Hour,
  train.Tuesday=reg.test$Tuesday,
  train.Wednesday=reg.test$Wednesday,
  train.Thursday=reg.test$Thursday,
  train.Friday=reg.test$Friday,
  train.Saturday=reg.test$Saturday,
  train.January=reg.test$January,
  train.February=reg.test$February,
  train.March=reg.test$March,
  train.April=reg.test$April,
  train.May=reg.test$May,
  train.June=reg.test$June,
  train.July=reg.test$July,
  train.August=reg.test$August,
  train.September=reg.test$September,
  train.October=reg.test$October,
  train.November=reg.test$November))

accu=data.frame(accuracy(fcast,reg.test$Load_kW))
accu$MAPE
# add mape result
mape_result$regression_model=accu$MAPE
# MAPE 20.70072 22.42678

######################## Decompose Model+stl ######################## 

stlf.model=stlf(load_kw_y.train, method='arima',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,load_kw_y.test))

# add mape result
mape_result$decom_arima_model=accu$MAPE

# method="arima" MAPE Training 1.206612 Test 15.193803                        BEST BEST BEST

stlf.model=stlf(load_kw_y.train, method='naive',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,load_kw_y.test))

# add mape result
mape_result$decom_naive_model=accu$MAPE

# method="naive" MAPE Training 2.151739 Test 15.202198

stlf.model=stlf(load_kw_y.train, method='rwdrift',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,load_kw_y.test))

# add mape result
mape_result$decom_rwdrift_model=accu$MAPE

# method="rwdrift" MAPE Training 2.151706 Test 15.518230

######################## HoltWinters Model ########################
## Seasonal Holt-Winters
HW.model=HoltWinters(load_kw_y.train, seasonal = c("additive"))#"additive", 
HW_prediction=forecast(HW.model,h=ntest)
accu=data.frame(accuracy(HW_prediction,load_kw_y.test))

# add mape result
mape_result$HoltWinters_additive_model=accu$MAPE
# seasonal = c("additive") MAPE Training 2.501263 Test 18.033887

HW.model=HoltWinters(load_kw_y.train, seasonal = c("multiplicative"))
HW_prediction=forecast(HW.model,h=ntest)
accu=data.frame(accuracy(HW_prediction,load_kw_y.test))

# add mape result
mape_result$HoltWinters_multiplicative_model=accu$MAPE
# seasonal = c("multiplicative") MAPE Training 2.344205 Test 18.287523


## Exponential Smoothing
HW_expo.model <- HoltWinters(load_kw_y.train, gamma = FALSE, beta = FALSE)
HW_expo_prediction=forecast(HW_expo.model,h=ntest)
accu=data.frame(accuracy(HW_expo_prediction,load_kw_y.test))

# add mape result
mape_result$HoltWinters_expo_model=accu$MAPE
#  MAPE Training 7.078471 Test 29.019098

print(mape_result)


# decom_arima_model is the best
## predict for 2005 and save to "Load Prediction 2005.xlsx"
ntest2005=8760

stlf.model=stlf(load_kw_y, method='arima',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest2005)

DailyPeakHour <- read_excel("Competition data.xlsx", sheet="DailyPeakHour")

ph=c(na.omit(DailyPeakHour$Load_kW),stlf_prediction$mean)
DailyPeakHour$Load_kW=ph
predict_2005=select(DailyPeakHour,Date, Hour,Load_kW)

write.xlsx(predict_2005, "Load Prediction 2005.xlsx",sheetName="Prediction")



# get DailyPeakHour and save to "DailyPeakHour.csv"
#Load_kW_regression
temp=Load_kW_regression%>% 
  group_by(Date)%>%
  summarise(
    highest_power=max(Load_kW)
  )
test=merge(temp, Load_kW_regression, by.x=c("Date", "highest_power"), 
           by.y=c("Date", "Load_kW"),
           all.x=TRUE)
lengths(unique(test["Date"]))==lengths(unique(temp["Date"]))
DailyPeakHour=test

DailyPeakHour=data.frame(DailyPeakHour)
DailyPeakHour$trend=1:length(DailyPeakHour$Date)

drops=c("T","Time","DayofWeek..1..Sunday..to.7..Saturday..","Datetime")

tt=DailyPeakHour[ , !(names(DailyPeakHour) %in% drops)]

write.xlsx(tt, file = "DailyPeakHour.xlsx")

