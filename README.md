# DataProcessing
This repo containts two types of information: i) Theoretical guides that help organize the data-processing approach, and ii) scripts in R and Excel VBA that add data-processing functionality. This repo is a Work-In-Progress... I will be adding another 20 scripts in the next few months.

## R scripts
Some scripts for processing of data. Eventually there will be a thorough map of scripts, for now it's just random thingies here and there.


### CompareStrings.R
Compares one list of strings against a list of correct strings and proposes matches based on string distances.

### misspellingFixer.R
Takes a long character vector with duplicate entries, some with misspellings and tries it's best to identify the misspellings and propose the fixes.


## Macros in Excel VBA
Collection of macros and functions that assist in the data cleaning, manipulating, and analyzing


### make_cumulative() example

Given:

| Item | 1  | 2 | 3  | 4 | 5   | 6 | 7 | 8 | 9 | 
|------|----|---|----|---|-----|---|---|---|---| 
| **A**    | 55 |   | 3  |   |     | 1 |   | 5 |   | 
| **B**    |    | 1 |    |   | 0.5 |   |   |   |   | 
| **C**    | 11 |   | 11 |   | 115 |   |   | 6 |   | 

We want to figure out how the items accumulated over the course of time... so select "--------------->" (YES).
The result will become:

| Item | 1  | 2  | 3  | 4  | 5   | 6   | 7   | 8   | 9   | 
|------|----|----|----|----|-----|-----|-----|-----|-----| 
| **A**    | 55 | 55 | 58 | 58 | 58  | 59  | 59  | 64  | 64  | 
| **B**    | 0  | 1  | 1  | 1  | 1.5 | 1.5 | 1.5 | 1.5 | 1.5 | 
| **C**    | 11 | 11 | 22 | 22 | 137 | 137 | 137 | 143 | 143 | 

If we instead wanted the totals by year, then select down (NO).
Result will become:

| Item | 1  | 2 | 3  | 4 | 5     | 6 | 7 | 8  | 9 | 
|------|----|---|----|---|-------|---|---|----|---| 
| **A**    | 55 | 0 | 3  | 0 | 0     | 1 | 0 | 5  | 0 | 
| **B**    | 55 | 1 | 3  | 0 | 0.5   | 1 | 0 | 5  | 0 | 
| **C**    | 66 | 1 | 14 | 0 | 115.5 | 1 | 0 | 11 | 0 | 

### fillSeries() example
Given:

| a  | b        | c | d | e     | f    | g     | h    | i    |    j  | 
|--------|----------|---|---|-------|-------|-------|------|------|------| 
| **2**  | Hi       |   |   |       |       |       |      |      |      | 
| **3**  |          |   |   |       |       |       |      |      |      | 
| **4**  |          |   |   |       |       |       |      |      |      | 
| **5**  | hello    |   |   | three |       |       | four |      |      | 
| **6**  |          |   |   |       |       |       |      |      |      | 
| **7**  |          |   |   |       |       |       |      |      |      | 
| **8**  |          |   |   |       |       |       |      |      |      | 
| **9**  | fourteen |   |   |       |       |       |      |      |      | 
| **10** |          |   |   |       |       |       |      |      |      | 
| **11** | five     |   |   |       |       |       |      |      |      | 



This is the effect of running the macro twice, once on B2 (going down), and once on E5 (going right):

| a  | b        | c | d |   e   | f     | g     | h    | i    | j     | 
|----|----------|---|---|-------|-------|-------|------|------|------| 
| **2 ** | Hi       |   |   |       |       |       |      |      |      | 
| **3**  | Hi       |   |   |       |       |       |      |      |      | 
| **4**  | Hi       |   |   |       |       |       |      |      |      | 
| **5**  | hello    |   |   | three | three | three | four | four | four | 
| **6**  | hello    |   |   |       |       |       |      |      |      | 
| **7**  | hello    |   |   |       |       |       |      |      |      | 
| **8**  | hello    |   |   |       |       |       |      |      |      | 
| **9**  | fourteen |   |   |       |       |       |      |      |      | 
| **10** | fourteen |   |   |       |       |       |      |      |      | 
| **11** | five     |   |   |       |       |       |      |      |      | 

