# Format Multiple Sheets
[![JavaScript Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](https://standardjs.com)

## TL;DR

Use this script to copy everything except content  (formatting, col/row sizes ...) from your template sheet to the rest of sheets automatically in Google Spreadsheets.

## Why is this useful?

I've been in this situation: you want to make a big spreadsheet and every sheet needs to have the same format, column widths, row heights, etc. You usually start designing a template, then duplicate it to create the actual sheets. 

However, as time goes by, you may want to make some change to the design. You change the template and then, you have to go like this: 

`select range > copy template > change sheet > Menu > Edit > Paste special > Paste Format Only`, then **repeat** this process to every sheet, one by one. What if you have 30 sheets? 60? 100? (current limit)  *Nightmare!* 

But wait, what if you also need to match **the row heigths and column widths**? `... > Paste Format Only` just doesn't make the cut for you. The only way is going row by row, column by column, sheet by sheet, manually setting the sizes in pixels! Isn't it insane?

Not anymore :wink: , use this script to propagate the look from your template design sheet to the rest of your Spreadsheet. 

## Specifically, the script copies:

   - Formatting,
   - Col/rows sizes,
   - Tab color and
   - Frozen col/rows.
   - It also unhides any col/row that may be hidden (`0px`).

## To Do

   - Use a checklist to let the user decide what to copy/paste or not, i.e. the user may not want to overwrite the tab colors.
   - Give the option of non-destructive propagation of formatting. Maybe creating new sheets or a new Spreadsheet.
   - Refactor.
