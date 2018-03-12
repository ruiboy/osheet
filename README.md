# Oh Sheet!

Given a `in/in.csv` file of weekly OSH booking data, produces excel
spreadsheets:
* `out/am-booking.xlsx` - for staff to record children per am session
* `out/pm-booking.xlsx` - same for pm
* `out/am-signin.xlsx` - for sign-in/out by carers per am session
* `out/pm-signin.xlsx` - same for pm

## Config:
* `conf/term.csv` - the starting Monday for each term; used to compute term and week numbers
* `conf/allergy.csv` - configure some allergy flags against kids (beyond what is in the CSV)
* `conf/skip.csv` - don't add some kids to output

Note: kids names in these files are 'flipped' from the input file.
eg if input file has "Elmer J Fudd", this has "Fudd, Elmer J"

## Running:
`./bin/osheet.py --help` will show args and defaults

By default will generate 4 sheets for the upcoming Monday using above default
file locations, and will call this "Term X, Week Y" as per `conf/term.csv`

See `samples` for examples of all files.
