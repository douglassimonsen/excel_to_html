A tool for easily converting Excel sheets into equivalently formatted html.

## Use Case
My specific use case is sending automatic excel reports to management. Often, people will review these reports on their phone and complain (reasonably) about the difficulty of opening an Excel on their phone. Since other managers prefer Excels for their ability to do quick analyses on the fly, I cannot fully abandon the Excel attachment. This leads me to need to generate an excel attachment for a report and then manually recreate the summary statistics in an HTML table to stick into the body of the email. Given that many of these summary sections are identical to the summary sections found in the Excel, it seems reasonable to automate.

If you use Outlook and you'd like a simple tool to send automatic emails, check out another library of mine [here](https://github.com/mwhamilton/outlook_emailer)!


## Details
The program contains a single function designed for public consumption:
* main.main

## main.main
This function takes in the path to an Excel, a sheetname, and optional min/max row/column, and openpyxl_kwargs (passed to openpyxl.load_workbook)

```python
main(
  'test.xlsx',
  sheetname='Sheet1',
  min_row=0,
  max_row=3,
  min_col=1,
  max_col=None,
  openpyxl_kwargs={
      'data_only': True,  # converts formulas to their values
  }
)
```
