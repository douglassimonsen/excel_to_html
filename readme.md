A tool for easily converting Excel sheets into equivalently formatted html.

## Use Case
My specific use case is sending automatic excel reports to management. Often, people will review these reports on their phone and complain (reasonably) about the difficulty of opening an Excel on their phone. Since other managers prefer Excels for their ability to do quick analyses on the fly, I cannot fully abandon the Excel attachment. This leads me to need to generate an excel attachment for a report and then manually recreate the summary statistics in an HTML table to stick into the body of the email. Given that many of these summary sections are identical to the summary sections found in the Excel, it seems reasonable to automate.

If you use Outlook and you'd like a simple tool to send automatic emails, check out another library of mine [here](https://github.com/mwhamilton/outlook_emailer)!

## Choices made in Program
* Due to the fact that my primary target is the Microsoft Outlook client, I only use CSS1 styles. You can see available CSS in the outlook client [here](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338201(v=office.12))
* Similarly, in order to maximize portability and to simplify the first implementation, all styles are inlined. I expect to add the option to return a seperate CSS string and to reduce styles to classes.
* Edges cases for borders, such as a merged cell having multiple border-styles on a single edge are not replicated.
* I assume that the default borders (the light grey lines you see on a blank sheet) should be seen. Editing `static_values.DEFAULT_BORDER` can change it to be invisible.
* If a merged cell has a border that is outside of the viewing window, that border still appears.

## Details
The program contains a single function designed for public consumption:
* main.main

### main.main
This function takes in the path to an Excel, a sheetname, and optional min/max row/column, and openpyxl_kwargs (passed to openpyxl.load_workbook)

```python
main(
  'test.xlsx',
  sheetname='Sheet1',
  min_row=0,
  max_col=4,  # this will cut the merged cell in half!
  openpyxl_kwargs={
      'data_only': True,  # converts formulas to their values
  }
)
```
<img src="https://github.com/mwhamilton/excel_to_html/raw/master/excel_example.PNG" alt="Input Excel" width="45%"></img>
<img src="https://github.com/mwhamilton/excel_to_html/raw/master/html_example.PNG" alt="Ouput HTML" width="45%"></img>
