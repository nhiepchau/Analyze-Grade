library(tidyverse)
library(ggplot2)
library(readxl)
library(xlsx)
library(dplyr)
library(readr)
library(rmarkdown)

#Reading Excel Files into R
midGrade <- read_excel("E:\\R Project\\201_CO1007.xlsx", sheet = 2, range = "A5:AB162")
finalGrade <- read_excel("E:\\R Project\\201_CO1007.xlsx", sheet = 4, range = "A5:AO168")
glimpse(midGrade)
glimpse(finalGrade)
finalGrade[is.na(finalGrade)] = 0

#Computing How many true answer
midSeq <- rowSums(select(midGrade, `1`:`25`) == 1)
finalSeq <- rowSums(select(finalGrade, `1`:`38`) == 1)

#Computing Grade and Define Median - Max - Min
options(digits = 3)
midMark <- round((midSeq / 25) * 10, 2)
finalMark <- round((finalSeq / 38) * 10, 2)

median(midMark)
max(midMark)
min(midMark)

median(finalMark)
max(finalMark)
min(finalMark)

#Counting Student with conditions
sum(midMark >= 9)
sum(midMark >= 7)
sum(midMark >= 5)


sum(finalMark >= 9)
sum(finalMark >= 7)
sum(finalMark >= 5)
sum(midMark < 5)
sum(finalMark < 5)

#Define Group of Student (Highest and Lowest Grade)
midGrade$TO <- as.factor(midGrade$TO)
midGrade$MANH <- as.factor(midGrade$MANH)

midGrade <- mutate(midGrade, grade = midMark)
finalGrade <- mutate(finalGrade, grade = finalMark)

filter(select(midGrade, No, MANH, TO, grade), midGrade$grade == max(midMark))
filter(select(midGrade, No, MANH, TO, grade), midGrade$grade == min(midMark))
filter(select(finalGrade, No, MANH, TO, grade), finalGrade$grade == max(finalMark))
filter(select(finalGrade, No, MANH, TO, grade), finalGrade$grade == min(finalMark))

hist(midMark)
hist(finalMark)

hist(midMark, col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,60))
hist(finalMark, col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,50))

midPlot <- ggplot(midGrade, aes(x=midMark)) + 
  geom_histogram(color="#69b3a2", fill="#e9ecef",binwidth=1 ) + labs(x="diem", y="Tansuat", title="Pho diem Giua Ki")

#Define rank mark k_th
mark_K <- function(x, k)
{
    if (k == 1) 
    {  max(x)}
    else
    {  mark_K(x[x != max(x)], k - 1)}
}
amountMark_K <- function(x, k)
{
  sum(x == mark_K(x, k))
}
sort(midMark)

mark_K(midMark, 10)
amountMark_K(midMark, 10)

midPlot <- ggplot(midGrade, aes(x=grade)) + 
  geom_histogram(color="#e9ecef", fill="#69b3a2",binwidth=1) + labs(x="diem", y="Tansuat", title="Pho diem Giua Ki")
finalPlot <- ggplot(finalGrade, aes(x=grade)) + 
  geom_histogram(color="#e9ecef", fill="#69b3a2",binwidth=1) + labs(x="diem", y="Tansuat", title="Pho diem Cuoi Ki")
#render("sessionIII.Rmd", "html_document")

#Find group that exam well
finalGrade$MANH <- as.factor(finalGrade$MANH)
finalGrade$TO <- as.factor(finalGrade$TO)
levels(finalGrade$MANH)
levels(finalGrade$TO)

index_1A <- (finalGrade$MANH == "L01" & finalGrade$TO == "A")
average_1A <- mean(finalGrade$grade[index_1A])
average_1A
median(finalGrade$grade[index_1A])
h_1A <- hist(finalGrade$grade[index_1A], col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,10), xlim = c(0,10))
roundDens <- round(h_1A$density, 2)
data.frame(
   mocDiem = c("0 -> 1", "1 -> 2","2 -> 3","3 -> 4","4 -> 5", "5 -> 6","6 -> 7","7 -> 8" ,"8 -> 9"),
   soLuong = h_1A$counts,
   matDo = roundDens
)


index_1B <- (finalGrade$MANH == "L01" & finalGrade$TO == "B")
average_1B <- mean(finalGrade$grade[index_1B])
average_1B
median(finalGrade$grade[index_1B])
h_1B <- hist(finalGrade$grade[index_1B], col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,10), xlim = c(0,10))
roundDens <- round(h_1B$density, 2)
data.frame(
  mocDiem = c("0 -> 1", "1 -> 2","2 -> 3","3 -> 4","4 -> 5", "5 -> 6","6 -> 7","7 -> 8" ,"8 -> 9"),
  soLuong = h_1B$counts,
  matDo = roundDens
)

index_2A <- (finalGrade$MANH == "L02" & finalGrade$TO == "A")
average_2A <- mean(finalGrade$grade[index_2A])
average_2A
median(median(finalGrade$grade[index_2A]))
h_2A <- hist(finalGrade$grade[index_2A], col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,10), xlim = c(0,10))
roundDens <- round(h_2A$density, 2)
data.frame(
  mocDiem = c("0 -> 1", "1 -> 2","2 -> 3","3 -> 4","4 -> 5", "5 -> 6","6 -> 7","7 -> 8" ,"8 -> 9"),
  soLuong = h_2A$counts,
  matDo = roundDens
)

index_2B <- (finalGrade$MANH == "L02" & finalGrade$TO == "B")
average_2B <- mean(finalGrade$grade[index_2B])
average_2B
median(median(finalGrade$grade[index_2B]))
h_2B <- hist(finalGrade$grade[index_2B], col ="#69b3a2", right = FALSE, border = "#e9ecef", xlab = "diem", ylab = "so hoc sinh", ylim = c(0,10), xlim = c(0,10))
roundDens <- round(h_2B$density, 2)
data.frame(
  mocDiem = c("0 -> 1", "1 -> 2","2 -> 3","3 -> 4","4 -> 5", "5 -> 6","6 -> 7","7 -> 8" ,"8 -> 9", "9 -> 10"),
  soLuong = h_2B$counts,
  matDo = roundDens
)