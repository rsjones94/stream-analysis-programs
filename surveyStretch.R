rm(list=ls())
library(xlsx)


## get the data from the spreadsheet
fname = file.choose()

#fname = "X:\\Environmental Tools\\Survey Workup Tools\\Hardpoint_Worksheet.xlsx"

data = read.xlsx(fname, sheetIndex = 1, header = TRUE, stringsAsFactors = FALSE)

# separate the data into the static survey (datapull1) and the survey to be stretched (datapull2)
datapull1 = na.omit(data[,c(2,3)])
datapull2 = na.omit(data[,c(5,6)])

colnames = c("Station", "Elevation")
names(datapull1) = colnames
names(datapull2) = colnames

inRow = rbind(datapull1,datapull2)

xrange = c(min(inRow[,1]),max(inRow[,1]))
yrange = c(min(inRow[,2]),max(inRow[,2]))

#graphics

plot(NULL,xlim=xrange,ylim=yrange,xlab = "Station (ft)", ylab = "Elevation (ft)")

points(datapull1, col = 'red')
points(datapull2, col = 'blue')
lines(datapull1, col = 'red')
lines(datapull2, col = 'blue')

legendx = xrange[1]+(xrange[2]-xrange[1])*.6
legendy = yrange[1]+(yrange[2]-yrange[1])*.9

legend(legendx,legendy, c("Survey 1", "Survey 2"), lty = c(1,1), col = c("Red", "Blue"))

hardpointsSet = FALSE
numhardpoints = 0
firstSurveyHard = NULL
secondSurveyHard = NULL

print(quote=F, "Select hardpoints (Survey 1 first, then Survey 2). Hit esc to finish entering")
while(!hardpointsSet) {

  couplet = locator(2)$x
  
  if (!is.null(couplet)) {
    
    firstSurveyHard[numhardpoints+1] = couplet[1]
    secondSurveyHard[numhardpoints+1] = couplet[2]
    numhardpoints = numhardpoints+1
    abline(v = couplet, col = c("Red","Blue"))
    
  } else {
    
    if (numhardpoints < 2) {
      print(quote=F, "Minimum 2 hardpoints needed for adjustment")
    } else {
      hardpointsSet = TRUE
      print(quote=F, "Adjusting...")
    }
    
  }

}

# now find what real points the user input is nearest to

firstSurveySnapIndex = NULL
for (i in 1:length(firstSurveyHard)) {
  
  firstSurveySnapIndex[i] = which.min(abs(datapull1[,1] - firstSurveyHard[i]))
  
}

secondSurveySnapIndex = NULL
for (i in 1:length(secondSurveyHard)) {
  
  secondSurveySnapIndex[i] = which.min(abs(datapull2[,1] - secondSurveyHard[i]))
  
}


# get the snapped stationing
firstSurveySnap = sort(datapull1[firstSurveySnapIndex,1])
secondSurveySnap = sort(datapull2[secondSurveySnapIndex,1])

#display the snapped stationing

plot(NULL,xlim=xrange,ylim=yrange,xlab = "Station (ft)", ylab = "Elevation (ft)", main = "Snapped Hardpoints")

points(datapull1, col = 'red')
points(datapull2, col = 'blue')
lines(datapull1, col = 'red')
lines(datapull2, col = 'blue')

legendx = xrange[1]+(xrange[2]-xrange[1])*.6
legendy = yrange[1]+(yrange[2]-yrange[1])*.9

abline(v=firstSurveySnap, col = 'red')
abline(v=secondSurveySnap, col = 'blue')

legend(legendx,legendy, c("Survey 1", "Survey 2"), lty = c(1,1), col = c("Red", "Blue"))


#### adjust survey 2

# align the first snap points

adjustedSurvey = datapull2
adjustedSurveyX = adjustedSurvey[,1]
firstdiff = firstSurveySnap[1]-secondSurveySnap[1]
adjustedSurveyX = adjustedSurveyX + firstdiff

for (i in 1:(length(firstSurveySnapIndex)-1)) { # the linear stretch algorithm
  
  a = adjustedSurveyX[secondSurveySnapIndex[i]]
  b = adjustedSurveyX[secondSurveySnapIndex[i+1]]
  c = datapull1[firstSurveySnapIndex[i+1],1]
  
  gamma = c-b
  
  for (j in secondSurveySnapIndex[i]:length(adjustedSurveyX)) {
    
    p = adjustedSurveyX[j]
    lambda = (p-a)/(b-a)
    
    if (j > secondSurveySnapIndex[i+1]) {
      
      adjustedSurveyX[j] = adjustedSurveyX[j] + gamma
      
    } else {
      
      adjustedSurveyX[j] = adjustedSurveyX[j] + gamma*lambda
      
    }
    
  }
  
}


print(quote=F, "Completed")



adjustedSurvey[,1] = adjustedSurveyX

plot(NULL,xlim=xrange,ylim=yrange,xlab = "Station (ft)", ylab = "Elevation (ft)", main = "Survey Adjustments")

points(datapull1, col = 'red')
points(adjustedSurvey, col = 'blue')
lines(datapull1, col = 'red')
lines(adjustedSurvey, col = 'blue')

legendx = xrange[1]+(xrange[2]-xrange[1])*.6
legendy = yrange[1]+(yrange[2]-yrange[1])*.9

# these two sets of verticle ab lines should be the same, as the blue ones will end up snapping to the red ones
#abline(v=datapull1[firstSurveySnapIndex,1], col = 'red')
#abline(v=adjustedSurvey[secondSurveySnapIndex,1], col = 'blue')

abline(v=datapull1[firstSurveySnapIndex,1], col = 'mediumpurple')

legend(legendx,legendy, c("Survey 1", "Adjusted Survey"), lty = c(1,1), col = c("Red", "Blue"))



write.table(adjustedSurveyX, "clipboard", sep="\t", row.names=FALSE, col.names=FALSE)
print(quote=F, "Data copied to clipboard")






