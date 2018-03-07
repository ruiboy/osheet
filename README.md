# Oh Sheet!

Given a `in/in.csv` file of weekly care-booking data, produces excel
spreadsheets for printing for:
* `out/am-booking.xlsx` - staff's record of children per am session
* `out/pm-booking.xlsx` - same for pm
* `out/am-signin.xlsx` - sign-in/out by carers per am session
* `out/pm-signin.xlsx` - same for pm

## Config:
* `conf/allergy.csv` - configure some allergy flags against kids (beyond what is in the CSV)
* `conf/skip.csv` - don't add some kids to output

Note: kids names in this file are 'flipped' from the input file.
eg if input file has "Elmer J Fudd", this has "Fudd, Elmer J"

## Running:
`./bin/osheet.py --help` will show args and defaults

All input file locations are configurable.

See `samples` for examples of all files.
