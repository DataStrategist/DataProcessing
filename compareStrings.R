## This script compares one list of strings against a list of correct strings
## And proposes matches based on string distances.

library(dplyr)
library(tidyr)
library(stringdist)

## Read data in: Set working directory and 
## your Inputfile should have 2 columns: wrong strings, correct strings
setwd("path")
Input <- "file.csv"
a <- read.csv(Input,stringsAsFactors = FALSE)
b <- as.character(a[,2])
a <- as.character(a[,1])

## Uniquize, then convert to lower or Proper Case
a <- unique(a)
b <- unique(b)

## use this for lower
# a <- tolower(a)
# b <- tolower(b)

## Use This For Proper Case
simpleCap <- function(x) {
  s <- strsplit(x, " ")[[1]]
  paste(toupper(substring(s, 1,1)), substring(s, 2),
        sep="", collapse=" ")
}
a <- sapply(a, simpleCap)
b <- sapply(b, simpleCap)

## Make a matrix of stringdistances b/w each a and each b. Might wanna change the method
sd <- stringdistmatrix(a,b)
rownames(sd) <- a
colnames(sd) <- b

## Find the best match for each misspelling (if there's many matches, pick first)
result <- t(sapply(seq(nrow(sd)), function(i) {
  j <- which.min(sd[i,])
  c(rownames(sd)[i], colnames(sd)[j], sd[i,j])
}))
result <- result %>% as.data.frame 
result$V3 <- as.numeric(result$V3)

names(result) <- c("orig","candidate","dist")

## And investigate the results

result %>%
  arrange(dist) %>% select(dist) %>% table %>% as.data.frame %>% plot
result %>% arrange(dist) %>% View

## If you like it, write file!
result %>% 
  write.csv(paste(Input,"-output.csv",sep=""))
  