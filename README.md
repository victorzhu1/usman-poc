# Assignment: Serializing a chart in PPT/Excel add-ins

## Overview

You are given **2 Excel workbooks** with multiple worksheets and tables/charts. 3 charts will be **linked** into a PowerPoint presentation from various sheets/workbooks. Your task is to:
1. Make a button in the PPT add-in that triggers the serialization of the **currently selected chart**.
2. Find the underlying data range of **that specific chart** in the linked Excel sheet.
3. **Serialize that data** in some LLM-friendly way using the Excel add-in.
4. Send the serialization to an external API of your choice.

## Hard Requirements
- A functional PowerPoint add-in where the user can trigger the serialization.
- A supporting Excel add-in that can serialize the right chart.
- Evidence of external API recieving the serialization of that chart (just printing it out is sufficient).

## Optional/Bonus requirements
- Simulate a **code** response from the API to trigger a change on the chart's data. For example, make the chart display bars corresponding to years 1994, 1993, and 1991 with the values in descending order.
- Even better bonus: update the corresponding chart in PowerPoint with the chart you just changed in Excel
- Manage an **undo stack** to undo that change.

## Notes
- Don't worry about the appearance of the UI, we are concerned with the functionality.
- Make the code generalizable to other charts/PowerPoints, avoid hardcoding the functionality.
- A lot of the details are up to you, please document any key assumptions/notes somewhere.

Good luck!
