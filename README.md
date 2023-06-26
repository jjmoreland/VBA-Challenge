# VBA-Challenge
VBA Script for stock data

My VBA code was modeled after lessons learned in the ASU Bootcamp, June 25, 2023. Specifically, the lesson activity "Credit Card" provided the code to compare "stock" tickers and return the next ticker if the previous ticker was not equal. 
  If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

Our class instructor also provided helpful tips to calculate the values required for the "Additonal Functionality Section" (i.e Greatest % increase, % decrease, and volume). Specifically, the code noted below was given:
  If volume > greatest_volume Then
    greatest_volume = volume
    greatest_volume_ticker = ws.Cells(i, 1).Value
 End If
 This same idea was then replicated for printing greatest % increase and % decrease. For the greatest % decrease, the default value of 999999 was given so that values could be compared.
 (After much trial and error, imperative to have variables set to original values before the loop. Otherwise the "higher" values from previous sheets will carry over into subsequent sheets and thereby giving incorrect summary results).

 Pivot tables were created to compare original data to the VBA script values to ensure correct results were derived.

 Thanks
 Print("Justin.Moreland")
