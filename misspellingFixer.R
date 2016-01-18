## This script takes a long character vector with duplicate entries, some with
## misspellings and tries it's best to identify the misspellings and propose
## the fixes.

library(dplyr)
library(tidyr)
library(stringdist)
library(ggplot2)

## Read data
setwd("path")
Input <- "file.csv"
a <- read.csv(Input,stringsAsFactors = FALSE)
a <- as.character(a[,1])

## Try to identify misspellings. 
## If there's too many misspellings then maybe skip this section
a %>% table() %>% data.frame -> a.count
a.count %>% arrange() %>% ggplot(aes(x=Freq))+geom_bar()
kmeans(a.count$Freq,2, iter.max = 10, nstart = 5) -> k
plot(a.count$Freq, col = k$cluster, pch = 19, frame = FALSE, main = "K-means with k = 2")


## Make Leven algorithm matrix of all vs all
a <- unique(a)
sd <- stringdistmatrix(a,a)
rownames(sd) <- a
colnames(sd) <- a

## First eliminate half the matrix (since the top is the mirror of the bottom)
sd[upper.tri(sd)] <- NA

## Convert to flat
sd.df <- as.data.frame(sd)
sd.df$Name <- names(sd.df)
sd.flat <- gather(sd.df,Match,Distance,-Name)
sd.flat %>% filter(!is.na(Distance)&Distance !=0) ->sd.flat
  
## I am comparing the list to itself, so there are many zero length matches.
## Get rid of these, and start limiting to see how it is
sd.flat %>% 
  mutate(Name.length = stringr::str_length(Name)) %>% 
  mutate(Perc = round(Distance/Name.length*100,1)) -> sd.flat

## limit to the smallest of misspellings
sd.flat %>% 
  filter(Distance ==1) -> shortlist

shortlist %>% View

## write file!
shortlist %>% 
  write.csv(paste(Input,"-output.csv",sep=""))
  