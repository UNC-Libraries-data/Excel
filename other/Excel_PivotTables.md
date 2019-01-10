# Excel Lightning Talks: Pivot Tables

January 10th, 2019, Matt Jansen, [mtjansen@email.unc.edu](mailto:mtjansen@email.unc.edu)

## What can I use Pivot Tables for?

Pivot Tables help aggregate and summarize large datasets in a smaller table to help answer questions or reformat datasets.

### Example: 

You have an Excel spreadsheet with all of the research consultations with UNC Librarians last year.  Each row contains:
* date, description, 
* the patron's academic status (undergraduate, graduate, etc.) and department
* the librarian(s) involved

A PivotTable can help you answer questions like:

* Which departments did we help the most last year?
* What months or seasons were the busiest for a given librarian?
* How many consultations do we provide to undergraduates? to faculty?

All of these questions would be difficult and time-consuming to answer by *directly examining* a dataset with hundreds or thousands of rows, even using Excel formulas.  PivotTables provide an easy way to tabulate these elements of the dataset in easy to understand ways.

## Preparing data

Pivot Tables will be much easier to use with **clean**, machine readable tables.  This means:

* Column names are in the first row, and the data starts immediately in the second row (no extra spacing, multiline titles, merged cells etc.)
* Each row should represent "observations" at some common level.  Don't mix in groupings (e.g. All the states, and a US total in the same dataset).
* Each column or "Field" should be consistent and well-labeled.
* [Read more](https://www.jstatsoft.org/article/view/v059i10) about what constitutes clean or "Tidy Data"

## Demo - E Journal Data

This dataset covers 243 electronic journals and records their:

* Title
* LC Subclass
* Subject Team
* Subject
* 2016 Usage
* 2016 List Price
* 2016 Cost Per Use (CPU)
* 2017 List Price

**Note:**  I've changed this dataset to obscure true values, so don't make any judgements on any PivotTables we build off this data!

### Possible Questions:

* Which Subject area has the highest cost per use at UNC?
* How many journals do we subscribe to in each Subject Team?

## Learn More!

* [Research Hub Intro to PivotTables (with exercises and solutions)](https://unc-libraries-data.github.io/Excel/Excel_Workshop_Instructions#introduction-to-pivottables)
* Tableau, a tool for quickly building data visualizations, works in a very similar way to PivotTables.  Learn more at [Tableau I](https://calendar.lib.unc.edu/event/4874955?hs=a).
* [Make an appointment with Matt](https://guides.lib.unc.edu/mattjansen)