# Assignment: Serializing a chart in PPT/Excel add-ins

## Overview

You are given an Excel workbook with multiple worksheets and tables/charts. These charts will be **linked** into a PowerPoint presentation. Your task is to:
1. Trigger an event from the PowerPoint add-in from the slide with a Chart.
2. Find the underlying data range in the linked Excel sheet.
3. Serialize that data in an LLM-friendly way.
4. Send the serialization to an external API of your choice.

## Hard Requirements
- A functional PowerPoint add-in where the user can trigger the serialization.
- A supporting Excel add-in.
- 

## Optional/Bonus requirements
- Simulate a **code** response from the API to trigger a change on the chart's data.
- Manage an **undo stack** to undo that change.
  
