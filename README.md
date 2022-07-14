# stock-analysis
### Project Analysis
##### The  purpose of this project was to incorpoarte coding skills using Visual Basic Analysis (VBA) in relation to the VBA_challenge workbook within Excel. The VBA_challenge workbook consists of two dataset sheets that contain data on different stock types and there unique corresponding information. This information is seperated in two excel sheets; one contains the stock dataset for the year 2017 and the other sheet contains a stock dataset for the year 2018. Through VBA, a code is designed to create a new table that compiles and summarizes the data for each stock ticker by year. By doing this, we can obtain an output for the total daily volume and the return for each ticker, filtered by the year. This table is programmed to compare the percentage of return which will help the client make the best decision when deciding what stock is going to serve them the best. 
---
### Results
---
#### 2017 Stock Analysis
##### The 2017 Stock Analysis table was designed in VBA to loop through the different ticker rows within the 2017 Worksheet and compile the daily total volume from the volume column in the original sheet. In VBA, I was able to create these variables and assign an equation corresponding to the starting price and ending price column. The image below shows the code that specifies the variables presented in the table. Using through a series of if-then statements, I was able to color code the return percentages by thier positive and negative values. It is made clear in the 2017 table that only two stock types had a negative return. The two stock tickers were RUN and TERP highlighted in red; the rest of the tickers had a positive return which is highlighted in green. 
---
![image](https://user-images.githubusercontent.com/105329532/178869270-0a80811f-8394-492f-9964-4f9f5119d1d0.png)
![image](https://user-images.githubusercontent.com/105329532/178881402-ea809160-4b91-4be5-9314-1d2838ab128d.png)
---
#### 2018 Stock Analysis
##### The coding is also the same for 2018 within the format of the table and the outputs that the code targets. The only diffrence is that I was able to create an input specifically for looping through the data in 2018. By enetering the year 2018 in the message box shown in the image below, I was able to get a table of the stocks along with the variable data for the total daily volume and return percentage for 2018. As a result, the data shows us that in 2018 only two tickers had a positive return while the rest of the stocks that year had a negative return. Again, this is indicated by the color coding, green for positive and red for nagative. For both outputs, I was able to implement a run time messge box that allowed me to numerically capture the time of the macro run.
---
![image](https://user-images.githubusercontent.com/105329532/178883089-f1fd1f04-382d-42e3-ad03-7d071272aa09.png)

![image](https://user-images.githubusercontent.com/105329532/178882706-3431bd8d-8693-4635-bf72-f48275aa795a.png)
---
[image](https://user-images.githubusercontent.com/105329532/178868373-8b0f4e1f-dae9-4afe-aaa6-5677a4d6f82a.png)
---
## Summary
---
#### Advantages and Disadvantages of Refactoring Code
##### Advatages:
* ##### Refactoring code is advantageous in general because it allows your code to be moldable and adaptable. Having a structured code that you can make small changes and additions to allows one to make there code run better or in a more convenient matter without having to start from scratch. This allows one to build upon their work rather than change all of the aspects. It's very likely that the new additions that you make would be pretty similar in the foundation of your original code, so it is also time efficient and an effective way to recycle code. It allows the abillity to solve more problems with a strategy already in place.
---
##### Disadvantages:
* ##### Although there are many andvatages to refactoriing code, there are also downsides to doing this. I think that one of them is that you have to work within the constraints of your orginal code otherwise it could cause errors. Making more changes and additions to a code can also make it more error prone in general because you may be introducing lines of code that may not be compatible with the original outline.
---
#### Advantages and Disadvantages of Refactored VBA Script
##### Advantages:
* ##### Because a lot of the outline of the code and flow of inssructions was already completed in the original code, it made it easier for me to refactor the code with new information that I wanted to incorporate into the stock analysis table. I had more of an idea for where certain lines of code should go because a lot of the information repeated and carried the same nested for loop pattern that I could follow. Some of the VBA subscript outlines were multi-purposeful, so I was able to do two different things whil following the overall same script pattern.
---
##### Disadvantages: 
* ##### Similar to the disadvantages in coding in general, I faced errors when implementing some of my changes within the code. I had to be mindful of what lines of code I was adding and changing and how each of those changes effected the way the script ran. Keeping track of the changes as you go I think is good practice because it's easy to make a lot of changes and not know where the root of the error is within the code. VBA will only highlight where the issue is, but it will not give a lot of context on the root of the problem or solutions for it.
