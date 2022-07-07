# Stock-Analysis
Analyze Stocks using VBA

**Overview of Project:**  To find the feasibility of green stocks starting with DQ and then checking for varied stocks which were traded during the year 2017 and 2018. 

 *_Initialising :_* The stock DQ was analysed first over the years 2017 and 2018. 
 
 _Analysis :_ The DQ stock evvaluation also opened the doors to introduction of all the other available stocks, which were then added in the calculation sheet. 
The entire analysis was then refactored to reduce the time lapsed in the analysis. See below differences for both 2017 and 2018:

_(Initial code is available on module 1 and the refactored code is on module 4)_

<img width="218" alt="2017 - initial " src="https://user-images.githubusercontent.com/94858846/149639241-96e30b14-b344-4347-b8bb-c35e99d6c68f.png"><img width="245" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/94858846/149639253-75e3fa05-e78a-498e-b148-c9b1b95179c6.png">

<img width="222" alt="2018 - initial " src="https://user-images.githubusercontent.com/94858846/149639285-14c194dd-4317-4919-a728-2ce9c4372864.png"> <img width="239" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/94858846/149639292-1706c609-72c0-4f8f-90d3-7c25a6145612.png">

**Difference in code:** For the initial analysis the script was made that included two loops and the code runs for every ticker value. Refactoring the code so that the code populated the values as the row value changes, reduced the time taken by almost 5 times. 

Original code:

<img width="449" alt="image" src="https://user-images.githubusercontent.com/94858846/149639550-560475f8-0625-4a68-aac2-a33ac5a132ce.png"> 

Refactored code :

<img width="665" alt="image" src="https://user-images.githubusercontent.com/94858846/149639572-3a578e39-860d-41e1-b8f4-baa2e8ba8ca2.png">


**Results**

For the 2017, summary of stocks: 

<img width="240" alt="image" src="https://user-images.githubusercontent.com/94858846/149639664-45c0276b-13b0-45dc-b3b7-7a00b6837c03.png">


For the 2018, summary of stocks: 

<img width="238" alt="image" src="https://user-images.githubusercontent.com/94858846/149639642-e6e7c9dc-d699-4cea-84f0-09b558f9349e.png">


**Summary**

Refactoring made the code easier and faster to execute a code. 
Refactoring can be complex to script and might take more time to develop. 

Original code was easier to make but it took longer to execute. 
