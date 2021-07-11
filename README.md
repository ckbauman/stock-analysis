# Stock Analysis with VBA

## Overview of Project

Our client Steve has been reviewing Stock Analysis.  Originally he was only looking at one stock and expanded to 12 stocks.  He has been reviewing Daily Volume and Returns over the years of 2017 and 2018.

Total Daily volume is calculated by adding up the total purchases of each stock over the year.  Return is calculated by determing the stock prices at the beginning of the year and the end of the year and looking for a percent increase or decrease.

Now that we have developed the information for 12 specific stocks, it was determined that Steve may want to run an analysis on all stocks and not limit himself to 12.  This will require refactoring our analysis code to run more efficiently.  Specifically this will be done by transforming the code to run only one time, vs. loooping through multiple times for each index.

## Results

### Stock Comparison between 2017 and 2018

Stock analysis for 2017 shows that all but 1 of the stocks showed a positive percent increase on it's Return over the year.  

INSERT IMAGE

Specifically **DQ**, **ENPH** and **SEDG** had very large Returns.  You can see why Steve's parents decided they wanted to invest in **DQ** in 2018.

The only negative Return was for **TERP** at -7.2%.

Stock analysis for 2018 shows that most stocks showed a negative percent increase on it's Return over the year

INSERT IMAGE

**DQ** took a huge drop at -62.6% but **ENPH** and **RUN** were the only stocks that remained positive over the 2 years.  In general, 2018 stocks did not perform well


### Execution times between Original and Refactored script

We created 2 seperate alaysis macros.

The first was

'''
sub jjjjj
'''


The analysis was completed using the same initial dataset used in the prior analysis.

Following are the results of the data broken out to include only Plays and outcomes with their Goal amount.  Percentages were created in each category (Successful, Failed, Canceled).

The following chart includes the Percent Outcomes per Dollar amount Goals for Plays.

### Challenges and Difficulties Encountered

#### Challenges for Theater Outcomes by Launch Date

Please note that Theater data includes all information for Musicals, Plays and Spaces.  Plays make up approximatley 76% of the dataset, while Spaces is 13% and Musicals is at 10%.  Anlysis of this data broken out by subcategory show that Spaces and Musicals are very sporatic, where Plays determines the primary trend for Theater productions.  In hindsight, this dataset should only be looking at Plays to keep our analysis consistent between the 2 review categories.

#### Challenges for Outcomes Basaed on Goals

Both the Success and Failure lines are fairly consistent until around the $35,000 amount.  Further analysis should be done to determine what is causing these anomalies in the dataset.

Also note that the Dollar amounts were not converted based on currency.  This could explain some of the anomalies in the dataset.

## Results

### Outcomes based on Launch Date?

- The graph indicates that Theater launches are most successful in the Summer months of May, June and July with a declining rate as you head into the Fall and Winter.  Launch dates in the months of November and December see the lowest success rates and may be affected by the Holidays.

- Data also suggests that about 40% of launches fail to meet their goal.  Very few launches are canceled and have little impact on the data.

### Outcomes based on Goals?

- The graph indicates that Plays generally failed to meet their goal as the dollar amounts increased.

- The graph also indicates that Plays generally succeeded to meet their goal up to about $20,000 funding goals.

### Limitations of this dataset?

- The 2 reviews in the analysis should have been reviewing the same breakout to produce more consistency.  Either all Theaters or all Plays would have allowed us to standardize the analysis more efficiently.

- Other possible tables and/or graphs might include an analysis of length of the campaign vs. outcome as well as appropriately converting the correct currency amounts in the table.
