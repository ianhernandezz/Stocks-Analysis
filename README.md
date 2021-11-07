# For Steve - Stocks Analysis 
Creating a stocks analysis for our client Steve


**1.	Overview of Project: Explain the purpose of this analysis.**

  The purpose of the creation of this analysis is to collect information for our client, Steve. The analysis of this specific workbook is regarding stocks that   Steve is interested in. 


2017             |  2018
:-------------------------:|:-------------------------:
<img width="270" alt="VBA_Results_2017" src="https://user-images.githubusercontent.com/92615504/140557886-5238fa5a-51ef-4f52-be5e-889bd2cf1276.png">|  <img width="269" alt="VBA_Results_2018" src="https://user-images.githubusercontent.com/92615504/140557897-a4fd16ca-aec0-4bf1-9991-3c1ebe967193.png">

*the results presented above are based on the **refactored** code.


**The Original Code**

<img width="510" alt="Screen Shot 2021-11-06 at 9 28 34 PM" src="https://user-images.githubusercontent.com/92615504/140628837-2d3d18d7-3f8c-4426-9d66-7f8e9c67a986.png">

**Notes:**
- As it can be seen from this screenshot, it can be determined that this code is inneficient as it does have to run through all of the code one by one. 
- Doing one line at at a time causes the code to have to go through the process and come back to repeat it for each ticker on the array. 

**2.	Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.**

As the original code was being written and tested it can be determined from the execution of the code that it takes a bit of time to run the code.
2017             |  2018
:-------------------------:|:-------------------------:
|<img width="251" alt="VBA_Challenge_2017(1)" src="https://user-images.githubusercontent.com/92615504/140628489-1bd1bd5d-e9d4-4b7c-abf4-835ec5a14979.png">|<img width="253" alt="VBA_Challenge_2018(1)" src="https://user-images.githubusercontent.com/92615504/140628493-169310c0-37f6-4039-98be-a37152e58f1f.png">
 
With the refactored code it made a significant improvement from the original code as it can be seen from the execution for both years (2017 + 2018).


2017             |  2018
:-------------------------:|:-------------------------:
<img width="248" alt="VBA_Challenge2017" src="https://user-images.githubusercontent.com/92615504/140556763-f404915d-4047-47b0-8c46-72daefb8e357.png">| <img width="251" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92615504/140556760-e4022808-0006-47d2-8c10-2df25312d8d3.png">



**3.	Summary: In a summary statement, address the following questions.**

**a. What are the advantages or disadvantages of refactoring code?** 

 After the refactoring of the original code some of the notable advantages are; the timing and the organization. The original code, compared to the refactored,    was inefficient, which is why when the new code was execute, it delivered results much quicker. The organization was also a result of this, which made the code neater and exact.
    
**b. How do these pros and cons apply to refactoring the original VBA script?**    

  The application of the refactored code proves pros such as a more organized code, with descriptions as to what the code is meant to do. This allows for the ease of understanding of other people that might be taking a look at the code to understand it better at a glance. A con, which I personally ran into was, that when the code was being created, I had to really pay attention and follow the flow of the code if not I would get lost. Bringing it back to the pro of the descriptions so others will not get lost or misunderstand. 
    
    
