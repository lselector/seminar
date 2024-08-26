require(xts)
require(PerformanceAnalytics)
require(quantmod)

# xts demonstration
vec_length_three <- c(1, 2, 3)
vec_length_two <- c(1, 2)

this_fails <- cbind(vec_length_three, vec_length_two)

getSymbols('SPY', from = '1990-01-01')
getSymbols('TSLA', from = '1990-01-01')

head(TSLA)
head(SPY)

this_succeeds <- cbind(Ad(TSLA), Ad(SPY))
head(this_succeeds['2010-06-25::'])

getSymbols(c('SPY', 'TLT'), from = '1990-01-01')
returns <- na.omit(cbind(Return.calculate(Ad(SPY)),
                         Return.calculate(Ad(TLT))))

# another key feature of xts
returns['2008-01-01::2008-02-28']

first_strat <- Return.portfolio(R = returns, weights = c(.6, .4))
charts.PerformanceSummary(first_strat)
table.AnnualizedReturns(first_strat)
maxDrawdown(first_strat)
CalmarRatio(first_strat)


# get more assets

symbols <- c("SPY", "TLT", "IEF", "GLD", "DBC", "EEM", "EWJ", "EFA")
returns <- list()

# loop through symbol names
# append the adjusted return of each name to the list
for(symbol in symbols) {
  returns[[symbol]] <- Return.calculate(
    Ad(getSymbols(symbol, from = '1990-01-01', adjusted = TRUE, auto.assign = FALSE)))
}

# column bind returns, drop NAs
returns <- do.call(cbind, returns)
returns <- na.omit(returns)


# a very fundamental function to asset allocation strategies
ep <- endpoints(returns)
returns[ep,]

# breakdown of a basic asset allocation strategy  

weights <- list()

# how many months do we use to compute data
lookback = 6

# how many of the best assets do we want to hold
top_N = floor(ncol(returns)/2)

# create empty weight vector
tmp_weight <- data.frame(matrix(rep(0, ncol(returns)), nrow = 1, ncol = ncol(returns)))
colnames(tmp_weight) <- colnames(returns)

# loop through endpoints, using a subset of size <<lookback>>
for(i in 1:(length(ep)-lookback)) {
  
  # get your data subset
  subset <- returns[(ep[i]+1):(ep[i+lookback]),]
  
  # compute momentum 
  momentums <- Return.cumulative(subset)
  
  # rank momentums
  mom_ranks <- rank(-momentums)
  
  # select top assets
  selected_assets <- (mom_ranks <= top_N) & (momentums > 0)
  
  if(sum(selected_assets)==0) {
    weight <- tmp_weight
  } else {
    
    # subset on selected assets
    selected_subset <- subset[,selected_assets]
    
    # compute standard deviations for selected assets
    stdevs <- StdDev(selected_subset)
    
    # weigh using normalized inverse standard deviation
    inv_stdevs <- 1/stdevs
    weight <- inv_stdevs/sum(inv_stdevs)
    
    # as we only have a subset of assets but need to include assets we didn't select,
    # we do a bit of copying/assign by index trickery0
    empty_vec <- tmp_weight
    empty_vec[,colnames(weight)] <- weight
    weight <- empty_vec
  }
  
  # assign the date of the decision
  weight <- xts(weight, order.by=last(index(subset)))
  weights[[i]] <- weight
  
}

# combine the weights
weights <- do.call(rbind, weights)
strategy <- lag(Return.portfolio(R = lag(returns, -1), weights = weights, verbose = TRUE))

# analyze performance
charts.PerformanceSummary(strategy$returns)
table.AnnualizedReturns(strategy$returns)
maxDrawdown(strategy$returns)
CalmarRatio(strategy$returns)

# compute turnover
turnover <- abs(strategy$BOP.Weight - lag(strategy$EOP.Weight))
turnover <- xts(rowSums(turnover), order.by=index(turnover))
plot(turnover)