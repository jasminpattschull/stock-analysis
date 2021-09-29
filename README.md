# stock-analysis
Diversification of portfolio for Steve's parents
# Refactoring Stock Analysis Code For Efficiency


## Overview of Project: 
### Purpose:
The purpose of this analysis was to take code that had been provided to Steve and make it more efficient. When working through a large set of data clunky code can make it take longer to run. When providing an output for a client, you will always want to provide an efficient product.

### Background:
The product being provided to Steve is a stock analysis so he can easily compare stocks year over year. He is working for his clients and wants to provide them the best information for their investment purposes. The analysis compares the 12 stocks provided for years 2017 and 2018.  The total daily volume and return was calculated for each ticker. This most recent deliverable was to clean up the prior code provided and make it run more efficiently for Steve in the event that he wants to expand the dataset to include the entire stock market. This would be a substantial increase to the mere 12 stocks currently provided so the code needed to run more efficiently or risk losing valuable time as Steve waits for the code to run and give him the output.


## Results:
There were 12 tickers analyzed and they are the following:
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

The following loop was completed within the code to arrive at the output detailed below:
      Cells(4 + i, 1).Value = tickerIndex
      Cells(4 + i, 2).Value = tickerVolumes
      Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1

In 2017, it appears that 92% of the stocks had a favorable return.  DQ, ENPH, FSLR and SEDG had over 100% returns. TERP was the only stock that had a negative return. The total daily volumes calculated were substantial for both FSLR (at 1.4M)) and SPWR (at 1.6M). The stock with the lowest volume DQ, despite it having a return of 199.4%.

![2017 Analysis](https://user-images.githubusercontent.com/90632470/135357153-542cda45-d566-4e96-bcda-0e4570b40c4f.PNG)

In 2018, of the same 12 stocks compared in 2017, there was a drastic reduction in favorable returns and only 17% were favorable. ENPH and RUN were the only two in the red and also had a total daily volume of over 1M. SPWR also showed a total daily volume of over 1M though its return was negative at -44.6%. The two stocks with the lowest returns were DQ and JKS and both were on the lower end of the spectrum for total daily volume.

![2018 Analysis](https://user-images.githubusercontent.com/90632470/135357177-b619f19f-4fc3-4026-86a8-3dd2de552546.PNG)

The refactored code for 2017 took 2.18 seconds to run. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90632470/135357207-6bee97e4-2722-4149-910b-839919b80052.PNG)

The original code for 2017 took 2.04 seconds to run.

![2017 Timer](https://user-images.githubusercontent.com/90632470/135357304-d22fae75-5f8e-42e1-9f05-3ab747af7222.PNG)
![2017 Format Timer](https://user-images.githubusercontent.com/90632470/135357272-63aca8a7-eaee-4019-a8ca-6fc07c45c815.PNG)

The refactored code for 2018 took 2.25 seconds to run.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90632470/135357239-ddb19115-8414-4dd9-9a91-9eb9594b4568.PNG)

The originl code for 2018 took 1.90 seconds to run.

![2018 Timer](https://user-images.githubusercontent.com/90632470/135357343-d613a240-81a5-4995-8b22-76a5c817c082.PNG)
![2018 Format Timer](https://user-images.githubusercontent.com/90632470/135357361-e1d9cbe5-fcaf-41af-a71b-038d25d62eda.PNG)


## Summary: 
### Advantages/Disadvantages of Refactoring Code:
There are numerous advantages to refactoring code. I find efficiency to be of utmost importance, generally speaking, and when code is refactored it is done for more efficiency.  Increased efficiency means less strain on the system and faster results for the end user. Another benefit to refactored code is the ability to more easily find errors or debug as changes or updates are made. When there are less efficient means used in coding, this means more places for errors to hide so refactoring minimizes your potential for errors by reducing the places for them to be.

A disadvantage to refactoring code could potentially be spending time "fixing" something that doesn't necessarily need fixed. If the code isn't going to need any future edits and runs well, it may potentially be best to leave as is. When refactoring it's best practice to copy the code before making edits so a record is maintained prior to those changes as well. If this is forgotten and the code is broken, it may result in further wasted effort. 

### How Advantages/Disadvantages Apply to Refactoring the Original VBA Script?
Efficiency and cleaner code is likely never going to be a negative. With the fact that I'm new to VBA, it did take longer and likely would have taken someone more experienced less time to make the updates. My original code (ran as two separate subroutines; one for the analysis and one for the formatting) did take less time than the refactored code. I'm unsure if this was due to my system and the other windows/programs that may have been running or not. The original code does produce the correct results, though we learned in class that sometimes code can still have issues despite arriving at the correct answer. With that said, I'd like to reiterate that efficiency and refactoring of code is a benefit that should always be considered, especially when a client is involved and it means the code can run faster with less strain on the client's system.
