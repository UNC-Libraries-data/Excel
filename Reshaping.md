# Pivoting Using Power Query Editor

Sometimes we need to change the shape of a dataset by “pivoting” or “unpivoting” columns. Why would this be useful? GIS programs can be very particular about how certain tabular data is formatted. For example, you need to organize x and y coordinates in separate columns to use ArcMap’s “Display XY Data” function. Or, you need to organize your unique identifiers in a distinct column to perform a join. We will go over this second example here.

## Pivoting Columns

Take the following dataset. Here we have the output of four factories that have the unique IDs 101-104, for multiple years.

![Unpivoted data table](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Origin_table_pivot.PNG)

We have points representing the factories in ArcMap that we’d like to join this data to. What we need is for each factory ID to be in its own distinct row, with columns for each production year.

![Pivoting example](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Pivot_arrow.PNG)


*Note: These instructions cover pivoting using Windows Excel’s Power Query. This process can also be performed using Excel pivot tables, which is detailed in “Working with Data in Excel”*

1. In Windows Excel, select the “Data” tab -> Select “From Table/Range”
2. If "Create Table" prompt appears
   1. Drag to highlight your entire table
   2. Select "My table has headers"
   3. Select "Ok"

![Create table for Power Query](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Create_table_pivot.PNG)

3. Power Query Editor will open
4. Highlight the column you want to pivot. The attributes listed in the column you choose to pivot will become their own separate columns (in this example, “Year”).
5. Select the "Transform" tab -> Select "Pivot Column"
6. In the "Values Column" dropdown, select the column containing the data points you would like to fill under your new columns. In this example, we would like “Output” to fill in under “Year”.

![Pivot column settings](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Pivot_column.PNG)

7. Your data will pivot. Close the window and "Keep" your changes.

*Note: If you would like to undo any steps of the pivoting process, go to the “Applied Steps” window under the “Query Settings” pane on the right side of the screen, and click the “x”. If Query Settings is not open, click the “View” tab, and “Query Settings”.*

## Unpivoting Columns

What if you would like to unpivot columns? For example, you have the dataset below, showing the population of multiple cities over multiple years.

![Pivoted data table](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Origin_table_unpivot.PNG)

You’re hoping to calculate the average population of each city over the years in a pivot table. In order to do so, you need to structure your data so that city, population, and year are in different columns.

![Unpivoting examle](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Unpivot_arrow.png)

1. In Windows Excel, select the "Data" tab -> Select "From Table/Range"
2. If "Create Table" prompt appears
   1. Drag to highlight your entire table
   2. Select "My table has headers"
   3. Click "Ok"

![Create table for Power Query](https://github.com/UNC-Libraries-data/Excel/blob/master/media/Create_table_unpivot.png)

3. Power Query Editor will open
4. Select the columns that you would like to condense into a single column. In this example, we want Raleigh, Durham, and Chapel Hill to become a single “City” column. To select multiple columns, hold down the shift key.
5. Select the "Transform" tab -> Select "Unpivot Columns". Your data will unpivot.
6. Your new columns will be renamed “Attribute” and “Value”. Double click the column header to change the name, in this case, to “City”, and “Population”.
7. Close the window and "Keep" your changes
