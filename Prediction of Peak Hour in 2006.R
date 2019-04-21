library("readxl")
library(rvest)
library(dplyr)
library(forecast)
library(ggplot2)
library("openxlsx")

# get DailyPeakHour data from excel file
DailyPeakHour <- read_excel("DailyPeakHour.xlsx")

# create a time series object 
dph_y=msts(DailyPeakHour$Hour,seasonal.periods = c(7,30,365),ts.frequency=365,
             start = as.POSIXct("2002-01-01"))

# set the last 6 months as test, set previous months as traning 6*30*24=4320 points as test
#testing set = last 4320 points
ntest=6*30
ntrain=length(dph_y) - ntest

train.end=time(dph_y)[ntrain]
test.start=time(dph_y)[ntrain+1]

dph_y.train=window(dph_y,end=train.end)
dph_y.test=window(dph_y,start=test.start)



######################## Build a seasonal naive model ########################

naive_M1=snaive(dph_y.train,h=length(dph_y.test),level=95)

#autoplot(naive_M1) + autolayer(naive_M1$fitted,series="Fitted\nvalues")+
 # autolayer(dph_y.test,series="Testing\nset")   

accu=data.frame(accuracy(na.omit(naive_M1),dph_y.test))


# create mape dataframe
mape_result=data.frame("naive_model"=accu$MAPE,row.names = c("test","train"))
mape_result


######################## Smoothing model ######################## Exponential smoothing state space model

#M3=stlf(dph_y.train, h = ntest,lambda="auto",level=95,s.window = 30)#, 
#M3F=forecast(M3,h=ntest)

#autoplot(M3$fitted,series="Fitted\nvalues")+
 # autolayer(dph_y.test,series="Testing\nset")#+autolayer(M3F)

#accuracy(M3F,dph_y.test)


# ME     RMSE       MAE         MPE      MAPE      MASE      ACF1
# Training set   -440.1741 14338.38  7954.419  -0.1075418  2.352408 0.1362683 0.6839599
# Test set     -52054.5414 93195.60 70244.017 -18.7794022 23.428768 1.2033600 0.9589556


# Residual diagnostics
#tsdisplay(M3$residuals)
#Box.test(M3$residuals,lag = 24)
#checkresiduals(M3$residuals)


######################## ARIMA model ######################## 

Acf(dph_y.train,lag.max = 24,main="")

ari.model=Arima(dph_y.train,order=c(0,0,0),lambda="auto")
arima_prediction=forecast(ari.model,h=ntest)
accu=data.frame(accuracy(arima_prediction,dph_y.test))

accu$MAPE

# add mape result
mape_result$arima_model=accu$MAPE
mape_result

#autoplot(M4F) + autolayer(M4$fitted,series="Fitted\nvalues")+
#  autolayer(dph_y.test,series="Testing\nset")

######################## TBATS Model ######################## Exponential Smoothing State Space Model With Box-Cox Transformation, ARMA Errors, Trend And Seasonal Components

#?tbats
tba.model=tbats(dph_y.train,use.box.cox=TRUE,use.trend = TRUE,
                use.arma.errors=TRUE,use.parallel=TRUE, num.cores=NULL)

tba_prediction=forecast(tba.model,h=ntest)
accu=data.frame(accuracy(tba_prediction,dph_y.test))



# add mape result
mape_result$tbats_model=accu$MAPE

# MAPE Training 2.60807 Test 3585.08863

######################## STL Model ######################## applying a non-seasonal forecasting method to the seasonally adjusted data and re-seasonalizing using the last year of the seasonal component.


#The first letter denotes the error type (“A”, “M” or “Z”); 
#the second letter denotes the trend type (“N”,”A”,”M” or “Z”); 
#and the third letter denotes the season type (“N”,”A”,”M” or “Z”). 
#In all cases, “N”=none, “A”=additive, “M”=multiplicative and “Z”=automatically selected. 
# So, for example, “ANN” is simple exponential smoothing with additive errors, 
#“MAM” is multiplicative Holt-Winters’ method with multiplicative errors, and so on.


#?stlm
stlm.model=stlm(dph_y.train, s.window = "periodic", robust = FALSE, method = c("arima"),
                modelfunction = NULL, model = NULL, etsmodel = "MAM",
                lambda = NULL, biasadj = FALSE, xreg = NULL,
                allow.multiplicative.trend = FALSE)

stlm_prediction=forecast(stlm.model,h=ntest)
accu=data.frame(accuracy(stlm_prediction,dph_y.test))
accu$MAPE

# add mape result
mape_result$stl_model=accu$MAPE

# MAPE 15.86584 17.06902

######################## Regression Model ######################## 
#?tslm
DailyPeakHour

ntest=6*30
ntrain=length(DailyPeakHour) - ntest

reg.train=head(DailyPeakHour,ntrain)
reg.test=tail(DailyPeakHour,ntest)

reg.train.y=msts(reg.train$Hour,seasonal.periods = c(7,30,365),ts.frequency=365,
                                  start = as.POSIXct("2002-01-01"))

train.Monday=reg.train$Monday
train.Trend=reg.train$trend
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


train.lm.trend.season=tslm(reg.train.y~train.Trend+train.Monday+train.Tuesday+
                             train.Wednesday+train.Thursday+train.Friday+train.Saturday+
                              train.January+train.February+train.March+train.April+train.May+
                             train.June+train.July+train.August+train.September+train.October+train.November
                             ,lambda="auto")
summary(train.lm.trend.season)

fcast = forecast(train.lm.trend.season, newdata = data.frame(
  train.Monday=reg.test$Monday,
  train.Trend=reg.test$trend,
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
  train.November=reg.test$November
  ))

accu=data.frame(accuracy(fcast,reg.test$Hour))
accu$MAPE
# add mape result
mape_result$regression_model=accu$MAPE
# MAPE 25.18617 18.17361

######################## Decompose Model+stl ######################## 

# method="arima" 17.51580 16.88061
stlf.model=stlf(dph_y.train, method='arima',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,dph_y.test))

# add mape result
mape_result$decom_arima_model=accu$MAPE

                      
# method="naive" MAPE 23.28754 16.91098
stlf.model=stlf(dph_y.train, method='naive',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,dph_y.test))
# add mape result
mape_result$decom_naive_model=accu$MAPE


# method="rwdrift" MAPE 23.28947 17.09439
stlf.model=stlf(dph_y.train, method='rwdrift',lambda="auto",biasadj=FALSE)
stlf_prediction=forecast(stlf.model,h=ntest)
accu=data.frame(accuracy(stlf_prediction,dph_y.test))

# add mape result
mape_result$decom_rwdrift_model=accu$MAPE

######################## HoltWinters Model ########################
## Seasonal Holt-Winters
HW.model=HoltWinters(dph_y.train, seasonal = c("additive"))#"additive", 
HW_prediction=forecast(HW.model,h=ntest)
accu=data.frame(accuracy(HW_prediction,dph_y.test))

# add mape result
mape_result$HoltWinters_additive_model=accu$MAPE
# seasonal = c("additive") MAPE 18.31141 16.57884

HW.model=HoltWinters(dph_y.train, seasonal = c("multiplicative"))
HW_prediction=forecast(HW.model,h=ntest)
accu=data.frame(accuracy(HW_prediction,dph_y.test))
accu$MAPE
# add mape result
mape_result$HoltWinters_multiplicative_model=accu$MAPE
# seasonal = c("multiplicative") MAPE 18.32123 16.62039


## Exponential Smoothing
HW_expo.model <- HoltWinters(dph_y.train, gamma = FALSE, beta = FALSE)
HW_expo_prediction=forecast(HW_expo.model,h=ntest)
accu=data.frame(accuracy(HW_expo_prediction,dph_y.test))
accu$MAPE
# add mape result
mape_result$HoltWinters_expo_model=accu$MAPE
#  MAPE Training 23.40537 20.62065

print(mape_result)


# HoltWinters_additive_model(16.57884) is the best
## predict for 2006
ntest2006=365

HW.model=HoltWinters(dph_y, seasonal = c("additive"))
HW_prediction=forecast(HW.model,h=ntest2006)


predict_2006 <- read.xlsx("Competition data.xlsx", sheet="2006")
predict_2006$DailyPeakHour=HW_prediction$mean

write.xlsx(predict_2006, "DailyPeakHour Prediction 2006.xlsx")
