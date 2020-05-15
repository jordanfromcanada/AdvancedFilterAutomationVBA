# AdvancedFilterAutomationVBA
Perform multiple Advanced Filter actions dynamically in Excel using VBA. Useful for visualizing data in Excel to be saved as a PDF for display purposes and sharing.

In the example below, a manager has a table of employees and wants to visualize how their work is distributed across three projects.

# Example
An Excel workbook contains two worksheets:
![sheetnames](https://user-images.githubusercontent.com/65370643/82013672-130a5280-9638-11ea-8604-92419694ea9e.JPG)

Table data describing employees in Sheets("employee-data")
![table-data-list-for-filterint](https://user-images.githubusercontent.com/65370643/82013673-13a2e900-9638-11ea-99f6-7098a9768382.JPG)

Each project is linked to a **ProjectID**, e.g. '9999'.
The button 'Insert salaries' on the far right will append each employee's name, job title and salary underneath whichever project they are linked to in the "employee-data" table.<br/>
*This is equivalent to performing the Data > Filter > Advanced function three times with three separate sets of criteria (for each of the three projects).*
![sheet-before-filter](https://user-images.githubusercontent.com/65370643/82013678-156cac80-9638-11ea-811f-b5e8f857ce07.JPG)

After running the VBA, employee data is inserted under each project.
![sheet-after-filter](https://user-images.githubusercontent.com/65370643/82013676-14d41600-9638-11ea-80f8-a848600ace41.JPG)
