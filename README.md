# Broad-Street-QA-Automation

Note: This is repo is archieved and read-only since it's no longer in use.

 This is a automation of Quality Assurance track for Mountains in BroadStreet.io. Run the main program on the bottom. Program will ask you the state you want to work on and then it will help you find the date you can work on. 

## Main.ipynb
 This is main program file. Notice that one may need a credential from console.google.com and datasets on your hand for this to work.

 One thing that needs to be done is commenting on cells.

# Authentication

Please refer to [gspread documention](https://gspread.readthedocs.io/en/latest/oauth2.html#enable-api-access-for-a-project) and use OAuth Client ID.

# How to run it

- Run through `Main.ipynb` until it says `Main Program`.
- Change the state and date to your will.
- Keep running until it says `save your csv here`
- Open `.xlsx` file of your corresponding state and save it as `.csv` (Please don't change the name)
  - You can also add your state and `.xlsx` files to the `comparison sheets` folder
- Keep running the rest of the code, and copy the final cell's output to STATE TOTAL
- Iterate the process for multiple dates and states.

