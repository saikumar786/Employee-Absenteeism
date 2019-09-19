#First remove the data in environment
rm(list = ls())

#Loading the library which is useful for reading xls file(As we have provided data in xls format)
library(readxl)

#Load data
absent_data <- read_xls('Absenteeism_at_work_Project.xls', sheet = 1, col_names = TRUE)

#Making the data we loaded as a data frame
df <- as.data.frame(absent_data)
#Make a copy of the df                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
df_copy <- df

#Checking the head of the dataframe
head(df)

#Check the shape of data
dim(df) #740-Observations and 21-Variables including target

#Structure of the data(To know about the type of each variable,feature)
str(df)

#Features of the df
colnames(df)


#--------------------------Unique Value Count-----------------------------------------------#

#An empty list
list <- c()

for(feature in seq(1, ncol(df), by=1)){
  list[[feature]] <- length(unique(df[[feature]]))
  #print(df[[feature]])
}

#Making the above list as a nice readable format of dataframe
uniq_df <- data.frame(Unique_count = list, row.names = colnames(df))

#-------------------------------------------------------------------------------------------#

###############################-MISSING VALUE ANALYSIS-######################################
#Checking for missing values
missing_data <- data.frame(Missing_count = (apply(df, 2, function(feature){sum(is.na(feature))})))
print(missing_data)

#Renaming the column(Missing_count to Missing_percentage)
names(missing_data)[1] <- "Missing_Percentage"
#Now, calculate the Percentage of Missing values
missing_data$Missing_Percentage <- (missing_data$Missing_Percentage / dim(df)[1]) * 100
print(missing_data)

# We can observe that, Missing Percentage is < 30%. So, we go for Imputation

#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<-Imputation->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>#

#First, we take an original value from the column (df$`Absenteeism time in hours`[74]) whose value=3
#Actual value = 3
#Mean = 6.9
#Median = 3
#KNN = 2

#Make the value we have taken as NULL
df$`Absenteeism time in hours`[74] = NaN
#Now, try Mean method to Impute the missing data which we have taken
df$`Absenteeism time in hours`[is.na(df$`Absenteeism time in hours`)] <- mean(df$`Absenteeism time in hours`, na.rm = T)
#Check how close the value is imputed by mean method
df$`Absenteeism time in hours`[74]

# Load data again and Make the value NaN again
df <- as.data.frame(absent_data)
df$`Absenteeism time in hours`[74] = NaN
#Now, try Median method to Impute the missing value
df$`Absenteeism time in hours`[is.na(df$`Absenteeism time in hours`)] <- median(df$`Absenteeism time in hours`, na.rm = T)
#Check that imputed value how close to the original value
df$`Absenteeism time in hours`[74]

#Again load data and make that value Null
df <- as.data.frame(absent_data)
df$`Absenteeism time in hours`[74] = NaN
#Load package required for KNN Imputation
library(DMwR)

df <- knnImputation(df, k = 3)
#Check how close the imputed value is to the Actual value
df$`Absenteeism time in hours`[74]

#___By observing Mean,Median,KNN imputation methods, we can select Median method for imputing
# As it results in very closest value to the Actual value.

#Now, check for Missing values
apply(df, 2, function(feature){sum(is.na(feature))})
#There are no missing values, As we have imputed data---

#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>#

#############################################################################################

#---------------@@@@@@@@@@@@@@@@@@@@@ Preprocessing ---------------@@@@@@@@@@@@@@@@@@@@@#

#Drop ID column(As there is no use of this feature/column)
df <- df[,-1]

#Now, group the `Reason for Absence` for our convenience to model the data

# Group 1: columns 1-14
# Group 2: columns 15-17
# Group 3: columns 18-21
# Group 4: columns 22-28
# We don't have 20 in the Reason for Absence column in the given data
df_after_dummy <- fastDummies::dummy_cols(df, select_columns = "Reason for absence")

head(df_after_dummy)

# We are getting 2 extra columns(20.62447, 21.44690). So, we delete those 2 columns
colnames(df_after_dummy)
df_after_dummy <- df_after_dummy[,-47]
colnames(df_after_dummy)
#Now, columns numbers has been changed. So check again the which column to be deleted and delete
df_after_dummy <- df_after_dummy[,-48]
#Check whether the 2 columns are deleted or not
colnames(df_after_dummy)

#Now, we can remove Reason for absence_0 column as it is not a specified reason in the list(Reasons for Absence) given to us

#Keep only Reason for absence columns to be grouped
df_reason_cols <- subset(df_after_dummy[,21:48])
colnames(df_reason_cols)

#Grouping from Reason 1-14
reason_1 <- df_reason_cols[,c("Reason for absence_1", "Reason for absence_2", "Reason for absence_3",
                              "Reason for absence_4", "Reason for absence_5", "Reason for absence_6",
                              "Reason for absence_7", "Reason for absence_8", "Reason for absence_9",
                              "Reason for absence_10", "Reason for absence_11", "Reason for absence_12",
                              "Reason for absence_13", "Reason for absence_14")]

#Grouping from Reason 15-17
reason_2 <- df_reason_cols[,c("Reason for absence_15", "Reason for absence_16",
                              "Reason for absence_17")]

#Grouping from Reason 18-21
reason_3 <- df_reason_cols[, c("Reason for absence_18", "Reason for absence_19", "Reason for absence_21")]

#Grouping from Reason 22-28
reason_4 <- df_reason_cols[, c("Reason for absence_22", "Reason for absence_23",
                               "Reason for absence_24", "Reason for absence_25",
                               "Reason for absence_26", "Reason for absence_27",
                               "Reason for absence_28")]

reason_type_1 <- apply(reason_1, 1, max)
reason_type_2 <- apply(reason_2, 1, max)
reason_type_3 <- apply(reason_3, 1, max)
reason_type_4 <- apply(reason_4, 1, max)


#Removing the columns(Reason) from df_after_dummy dataframe
df_after_dummy <- df_after_dummy[, -21:-47]
#Now, Remove `Reason for absence` column from df_after_dummy
df_after_dummy <- df_after_dummy[,-1]


#Adding new columns of Reason_1,Reason_2, Reason_3,Reason_4 to df_after_dummy dataframe
df_after_dummy$Reason_1 <- reason_type_1
df_after_dummy$Reason_2 <- reason_type_2
df_after_dummy$Reason_3 <- reason_type_3
df_after_dummy$Reason_4 <- reason_type_4

colnames(df_after_dummy)
#Removing unnecessary column(Reason for absence_16)
df_after_dummy <- df_after_dummy[,-20]


#Make a copy of df_after_dummy
df_dummy_copy <- df_after_dummy


#---------------- Outlier Analysis ---------------------#

# Now, separate continuous and categorical type features as separate lists
cont_var <- c("Transportation expense", "Distance from Residence to Work", "Service time",
              "Age", "Work load Average/day", "Hit target", "Pet", "Weight", "Height",
              "Body mass index", "Absenteeism time in hours")

cat_var <- c("Month of absence", "Day of the week", "Seasons", "Disciplinary failure",
             "Education", "Son", "Social drinker", "Social smoker", "Reason_1", 
             "Reason_2", "Reason_3", "Reason_4")

for(i in cont_var){
  #print(i)
  #Detecting outlier
  outlier_val <- df_after_dummy[,i][df_after_dummy[,i] %in% boxplot.stats(df_after_dummy[,i])$out]
  #Selecting only the observation which doesn't have outliers
  df_after_dummy <- df_after_dummy[which(!df_after_dummy[,i] %in% outlier_val),]
}

#Replace all outliers with NA & Impute them
for(i in cont_var){
  out_val <- df_after_dummy[,i][df_after_dummy[,i] %in% boxplot.stats(df_after_dummy[,i])$out]
  df_after_dummy[,i][df_after_dummy[,i] %in% out_val] <- NA
}

#Imputing the data now
df_after_dummy <- knnImputation(df_after_dummy, k=3)

#Plotting boxplot for all continuous variables
for(i in cont_var){
  boxplot(df_after_dummy$`Absenteeism time in hours` ~ df_after_dummy[,i] , data = df_after_dummy,
          xlab = i, ylab = 'Absenteeism in hours', col = c("blue", "green"))
}

################################### Exploratory Data Analysis #################################
#Plots
library(ggplot2)
#Absenteeism vs Transportation Expense
absent_transport <- ggplot(df_after_dummy, aes(x=`Transportation expense`, y=`Absenteeism time in hours`)) + geom_point() + geom_smooth(method = "lm")
absent_transport

#Absenteeism vs Month
abs_month <- ggplot(df_after_dummy, aes(x=`Month of absence`, y=`Absenteeism time in hours`)) +
  geom_point() + geom_smooth(method = "lm")
abs_month

#Absenteeism vs Seasons
abs_season <- ggplot(df_after_dummy, aes(x=`Seasons`, y=`Absenteeism time in hours`)) + 
  geom_point() + geom_smooth(method = "lm")
abs_season

#Absenteeism vs Age
abs_age <- ggplot(df_after_dummy, aes(x=`Age`, y=`Absenteeism time in hours`)) +
  geom_point() + geom_smooth(method = "lm")
abs_age

library(gridExtra)
#An empty list
plt <- list()
for(i in 1:ncol(df_after_dummy)){
  print(i)
  plt[[i]] <- ggplot(df_after_dummy, aes_string(x=df_after_dummy[,i], y=df_after_dummy$`Absenteeism time in hours`)) + geom_point() + geom_smooth(method = "lm")
}
do.call(grid.arrange, plt)



#Statistical test(ANOVA)
aov(df_after_dummy$`Absenteeism time in hours` ~ df_after_dummy$`Day of the week`, data=df_after_dummy)
aov(`Absenteeism time in hours` ~ `Month of absence`, data=df_after_dummy)
aov(`Absenteeism time in hours` ~ `Seasons`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Disciplinary failure`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Education`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Son`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Social drinker`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Social smoker`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Reason_1`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Reason_2`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Reason_3`, data = df_after_dummy)
aov(`Absenteeism time in hours` ~ `Reason_4`, data = df_after_dummy)

#------------------------------- Correlation Analysis -------------------------------------#
#Correlation plot
library(corrplot)

corr_df <- round(cor(df_after_dummy[, cont_var]), 2)
print(corr_df)
#Plotting Correlation plot
corrplot(corr_df, method = "number", tl.cex = 0.5)


#Using corrgram
install.packages("corrgram")
library(corrgram)
corrgram(df_after_dummy[, cont_var], order = F, upper.panel = panel.pie, text.panel = panel.txt,
         main="Correlation Plot")
#We can see that Body mass index and Weight are highly correlated. So, we can take any 1 of 
#those 2 variables to avoid Multicollinearity

#Update the cont_vars
cont_var <- c("Transportation expense", "Distance from Residence to Work", "Service time",
              "Age", "Work load Average/day", "Hit target", "Pet", "Weight", "Height",
              "Absenteeism time in hours")


#---------------------------------- Feature Scaling ---------------------------------#

#Normality check
hist(df_after_dummy$`Hit target`)
hist(df_after_dummy$`Distance from Residence to Work`)
hist(df_after_dummy$`Month of absence`)
hist(df_after_dummy$`Absenteeism time in hours`)

#We can say that data is not distributed uniformly. So, we Normalize data, else we Standardize
#Normalization
for(column in cont_var){
  df_after_dummy[, column] <- (df_after_dummy[,column] - min(df_after_dummy[,column])) / 
    (max(df_after_dummy[, column] - min(df_after_dummy[, column])))
}


#---------------------------------- Splitting data ------------------------------#
#load required packages for splitting data
require(caTools)

set.seed(65)
#Splitting data into 70% TRUE(Train) & 30% FALSE(Test)
sample1 <- sample.split(df_after_dummy, SplitRatio = 0.7)

#The rows which have TRUE will be in this train dataframe
train <- subset(df_after_dummy, sample1 == TRUE)

#Remaining 30% data which is not marked as TRUE will fall under this test data
test <- subset(df_after_dummy, sample1 == FALSE)

#---------------------- Model Building -----------------------------------#

#------------------------------Linear Regression----------------------------------------
lm_model <- lm(train$`Absenteeism time in hours` ~ (train$`Month of absence` + train$`Day of the week`
                                                    + train$Seasons + train$`Transportation expense`
                                                    + train$`Distance from Residence to Work`
                                                    + train$`Service time` + train$Age
                                                    + train$`Work load Average/day`
                                                    + train$`Hit target` + train$`Disciplinary failure`
                                                    + train$Education + train$Son + train$Pet 
                                                    + train$`Social drinker` + train$`Social smoker` 
                                                    + train$Weight + train$Reason_1 + train$Reason_2
                                                    + train$Reason_3 + train$Reason_4))
#Summary of the model
summary(lm_model)

#Predictions
pred1 <- predict(lm_model, newdata = train)

#Error calculation
error <- train[,'Absenteeism time in hours'] - pred1

#Root Mean Squared Error
rmse <- sqrt(mean(error^2))
paste("RMSE:",rmse)

#Significance values of predictors
library(knitr)
kable(anova(lm_model), booktabs=T)

#------------------------- Decision Tree ------------------------------------------
#Load CART Packages
library(rpart)
library(rpart.plot)

#CART Model
dt_model <- rpart(`Absenteeism time in hours` ~ `Month of absence` + `Day of the week` + 
                    `Seasons` + `Pet` + `Son` + `Social drinker` + `Social smoker` +
                    `Transportation expense` + `Age` + `Service time` + `Distance from Residence to Work` +
                    `Work load Average/day` + `Hit target` + `Disciplinary failure` +
                    `Education` + `Weight` + `Reason_1` + `Reason_2` + `Reason_3` + `Reason_4`, data = train)

#Plot the tree using prp command defined in rpart.plot package
prp(dt_model)

#Predictions
pred_dt <- predict(dt_model, newdata = test)
pred.sse <- sum((pred_dt - test$`Absenteeism time in hours`)^2)
print(pred.sse)

#------------------------------- Random Forest ---------------------------------#
library(randomForest)
set.seed(65)

rf <- randomForest(`Absenteeism time in hours` ~ (`Month of absence` + `Day of the week` + 
                     `Seasons` + `Pet` + `Son` + `Social drinker` + `Social smoker` +
                     `Transportation expense` + `Age` + `Service time` + `Distance from Residence to Work` +
                     `Work load Average/day` + `Hit target` + `Disciplinary failure` +
                     `Education` + `Weight` + `Reason_1` + `Reason_2` + `Reason_3` + `Reason_4`), data = train, ntree = 100)

library("inTrees")
#Extract rules from rf
treelist <- RF2List(rf)

#Prediction
rf_pred <- predict(rf, test[, 'Absenteeism time in hours'])
regr.eval(test[,'Absenteeism time in hours'], rf_pred, stats='rmse')

print(rf)

#We can say that, Linear Regression is fitted a little for this data




