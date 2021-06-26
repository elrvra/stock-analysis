
# VBA of Wall Street

## Overview of Project

### Background

Steve has just graduated with his Finance degree. His parents are so proud, that they are going to be his first clients. Steve’s parents are passionate about green energy. They believe that as fossil fuels get used up there will be more and more reliance on alternative energy production. There are many forms of green energy to invest in, including hydroelectricity, wind energy, geothermal energy, and bioenergy. However, Steve’s parents have not done much research and have instead decided to invest all their money in DAQO New Energy Corp (ticker: DQ), a company that makes silicon wafers for solar panels. 

Steve has promised to look into these stocks for his parents, and his parents want to particularly know how actively DQ was traded. One way to measure this is to calculate the yearly return for DQ. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, if you invested in DQ at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year. Steve has created an excel file of green energy stocks in addition to DQ's, and is asking for guidance.

### Purpose

In this challenge, I will edit, or refactor the Module 2 solution code that used to help Steve with his analysis to his parents of the 2017-2018 data. Then loop through all the data one time in order to collect the same information from the module. Then, I will determine whether refactoring the code successfuly made the VBA script run any faster. Finally, I will present a written analysis that explains my findings. 

## Results

### Compare the stock performance between 2017 and 2018 (with screen shots and code)
After analysis, we see that DQ (DAQO) dropped over 63% in 2018, despite it being up by 199% in 2017. Steve will definitely want to offer some better stocks to his parents, since DQ might not be the best option for Steve's parents to invest in. Additionally, it does seem that most of the 2017 was a good year in the energy sector (92% of stocks are in the green), whereas in 2018 is the opposite (83% of stocks are in the red). 

The stocks that remain consistent in their growth percentage from 2017 to 2018, were ENPH (Enphase Energy Inc), RUN (Sunrun Inc) and TERP (TerraForm Power Inc). Although remaining in the green YOY, ENPH's growth was higher in 2017 at 130% compared to 82% in 2018, and RUN increased significantly from 6% to 85%; TERP remained in the red, despite a lower negative growth rate. 

After refactoring the code, I can confirm that the stock analysis outputs for 2017 and 2018 are the same as the dataset we completed in the initial code earlier, and the screen shot of earlier analysis are as follows... 

![alt tag](https://github.com/elrvra/stock-analysis/blob/master/VBA_Challenege_Analyses_Before_And_After_Refactoring.png)

Additionally, the code that was required to be refactored, as per the Deliverable requirements and with explanation are as follows... 

![alt tag](https://github.com/elrvra/stock-analysis/blob/master/VBA_Challenege_Screenshot_of_Code.png)

### Compare the execution times from the original script and the refactored script
- For 2017, refactoring did cut down the execution time from the original script by -0.437500
- For 2018, refacoring increased the execution time from the original script by +0.031250
These changes can be shown in the tables as follows... 

![alt tag](https://github.com/elrvra/stock-analysis/blob/master/VBA_Challenege_Analyses_Before_And_After_Refactoring.png)

## Summary

### Details on the advantages and disadvantages of refactoring code in general
- Advantage: Refactoring makes the VBA code more efficient by breaking down an eliminating unncessary code
- Advantage: Refactoring improves the VBA logic making it more visual friendly for future users to read
- Disadvantage: Refactoring takes more time, so if there are projects with time constraints, it may not work
- Disadvantage: Refactoring can cause more bugs to sift through especially when there is a lot of code

### Details on advantages and disadvantages of the original and refactored VBA
- Advantage: Refactoring improved the execution time on the 2017 output
- Disadvantage: Refactoring added more seconds to the execution time on the 2018 output
- Advantage: Helped me to understand the code original completed better second time around
- Disadvantage: Still unsure as to why the 2018 data took longer to run than the original
