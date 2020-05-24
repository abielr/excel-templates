# Excel Templates

Excel Templates makes it easy to replace keys inside an Excel file with data from a Python dictionary. A typical use case would be to populate formatted tables in Excel without having to use the `VLOOKUP` function. It also makes it easy to tile a template within or across worksheets before populating it, all while maintaining formatting and updated formulas.

## Installation
```
pip install excel_templates
```

## Examples

### Basic usage
Suppose we've created a financial statement template in Excel, such as the one in the image below which shows quarterly revenue, expenses, and profits for a particular product and region combination.

![Blank Template](https://github.com/abielr/excel-templates/blob/master/doc/images/template_blank.png)

Notice that the Revenue and Expense cells of the table have been filled in with unique key that is the combination of the line item and period name. This can be created manually or as an Excel formula. The Profit line is an Excel formula that is equal to Revenue - Expense; it currently shows a `#VALUE!` error since the cells that it references have not been populated with numbers.

Let's start by filling just a couple values, namely the PRODUCT cell and Revenue for the first quarter of the year.

```python
from excel_templates import ExcelTemplate

data = {'Revenue/Q1': 100, 'PRODUCT': 'Apples'}
template = ExcelTemplate("template.xlsx")
template.fill("Sheet1", data=data)
template.save('output.xlsx')
```

Now `output.xlsx` will contain a table that looks like the following:

![Template1](https://github.com/abielr/excel-templates/blob/master/doc/images/template1.png)

Typically your raw data will already reside in a `pandas DataFrame` object. The package provides a helper function called `make_dict` to quickly concatenate columns in a `DataFrame` and translate it into a dictionary object. In the example below we simulate some data with multiple products and regions, then fill in the table with the data for a single product/region combination.

```python
from excel_templates import ExcelTemplate, make_dict
import pandas as pd
import numpy as np

np.random.seed(1)
full_data = []
for product in ['Apples','Bananas']:
    for region in ['Northeast','South']:
        for concept in ['Revenue','Expense']:
            df = pd.DataFrame({
                'product': product, 'region': region, 
                'period': ['Q1','Q2','Q3','Q4'],
                'concept': concept, 'value': np.random.rand(4)
            })
            full_data.append(df)
full_data = pd.concat(full_data)

data = full_data.query("product=='Apples' & region=='Northeast'")
data = make_dict(data, keys=['concept','period'], value='value', sep='/')
data['PRODUCT'] = 'Apples'
data['REGION'] = 'Northeast'

template = ExcelTemplate("template.xlsx")
template.fill("Sheet1", data=data)
template.save('output.xlsx')
```

This produces the following output

![Template2](https://github.com/abielr/excel-templates/blob/master/doc/images/template2.png) <br/><br/>

### Tiling templates
The simulated data above has two products and two regions, so let's tile the table into a 2x2 grid with products across the columns and regions down the rows. Continuing with the same simulated data as before:

```python
template = ExcelTemplate("template.xlsx")
template.tile('Sheet1', rows=2, columns=2)

for i, region in enumerate(full_data['region'].unique()):
    for j, product in enumerate(full_data['product'].unique()):
        data = full_data[(full_data['product']==product) & (full_data['region']==region)]
        data = make_dict(data, keys=['concept','period'], value='value', sep='/')
        data['PRODUCT'] = product
        data['REGION'] = region
        template.fill("Sheet1", data, i + 1, j + 1)

template.save("output.xlsx")

```

The first column of output is shown below, the second column would look similar but with the other product. By default there is one blank row and column between the tiled tables, this can be adjusted with the `row_spacing` and `col_spacing` arguments to `tile()`.

![Template3](https://github.com/abielr/excel-templates/blob/master/doc/images/template3.png) <br/><br/>

### Blank cells

Any cell whose value does not match a key in `data` is left unchanged. This may be not be what you want: for instance, continuing with the above example, suppose that Revenue and Expense for Apples in the Northeast was missing for Q4. Then the corresponding cells in the table would show `Revenue/Q4` and `Expense/Q4`. Instead, you may wish to fill those cells with another value, such as zero. This can be achieved by setting the `fillna` argument in `fill()`. However, you will usually want to use this in tandem with the `prefix` argument, which will only do substitution on cells whose values begins with the given prefix. This stops cells which are not supposed to be substituted, such as the row and column headers, from being erased.

The picture below shows this in action.

![Blank Template](https://github.com/abielr/excel-templates/doc/images/template_blank2.png)

Now the cells whose values which should substituted with values from `data` all begin with `//`. If we want to fill cells who do not match any key in `data` with zero we can write

```python
template.fill("Sheet1", data=data, fillna=0, prefix='//')
```

Note that the prefix, which in this case is `//`, is stripped from the Excel values before they are compared to keys in `data`, i.e. the value `//Revenue/Q4` in Excel will be substituted with the value whose key is `Revenue/Q4` in `data`. The image below shows the output when the Q4 data is missing.

![Filled Template](https://github.com/abielr/excel-templates/blob/master/doc/images/template_filled.png) <br/><br/>

### Copying worksheets

The `ExcelTemplate` class also has a `copy_worksheet` method which can be used to generate copies of worksheets, which can then be filled just like any other worksheet:

```python
# Copies Sheet1 to a new worksheet called Sheet2
template.copy_worksheet('Sheet1', 'Sheet2')
```

## Acknowledgments

* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)