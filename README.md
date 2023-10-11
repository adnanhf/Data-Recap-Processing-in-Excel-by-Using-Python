# Monthly Report Maker

## My Motive
I saw staff at my office doing a monthly recapitulation using Microsoft Excel manually. It takes around 3 to 5 days for one person to finish. Then I started to write this program, hoping to ease the pressure and speed up the time to complete the work.

## Disclaimer
This program's design would process raw data into a hard-coded output. If anyone would use this code, they will need to re-write the properties of both the input and output.

## Program's workflow

### A. File preparations
1. Manually input data to an `xlsx` file
2. Place the `xlsx` file inside `files` folder

### B. Running Sequence in main.py
1. Reads the `xlsx` file
2. Sorts data by date and time
3. Split data into categories
4. Re-create data into some column formats
5. Write sheets using newly generated data
6. Write load's resumes for some sheets
7. Program stopped
