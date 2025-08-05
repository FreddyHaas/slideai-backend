from typing import List

import pandas as pd

from app.models import ChartType


def create_chart_selection_prompt(df, chart_options: List[ChartType], core_message: str, header_cell_formats) -> str:
    # Extract column names and descriptions
    columns = df.columns.tolist()

    # Summarize data overview (e.g., range or unique values for each column)
    data_overview = []
    for col in columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            summary = f"Range: {df[col].min()} to {df[col].max()}"
        else:
            unique_vals = df[col].unique()
            unique_count = len(unique_vals)
            examples = ', '.join(map(str, unique_vals[:3]))  # Show up to 3 examples
            summary = f"Number of unique values: {unique_count} | Examples: {examples}"
        data_overview.append(f"- {col}: {summary}")
    data_overview_text = "\n".join(data_overview)

    # Include the first 5 rows of the DataFrame
    first_five_rows = df.head(5).to_string(index=False)

    # last 3 rows
    last_three_rows = df.tail(3).to_string(index=False)

    # Include header cell formats
    header_formats = "\n".join([f"- {col}: {fmt}" for col, fmt in header_cell_formats.items()])

    # Prepare chart options text
    chart_options_text = "\n".join(
        [
            f"*{option.value}*\n"
            f"   - Purpose: {option.purpose}\n"
            f"   - Required data: {option.data_input}"
            for option in chart_options
        ]
    )

    # Construct the prompt
    prompt = f"""
I have a table with the following summary characteristics:

- **Column names:**
{columns}

- **Data overview:**
{data_overview_text}

- **Column cell formats:**
{header_formats}

- **First 5 rows of the data:**
{first_five_rows}

- **The last 3 rows of the data:**
{last_three_rows}

The chart should support the following message:
"{core_message}"

**Task:** 

1. Based on this summary of the table, please explain your reasoning and 
select up to three appropriate chart types from the following options:
{chart_options_text}

Avoid selecting inappropriate chart types:

- If numeric values across columns are in different units, do not select clustered column charts, stacked column 
charts or 100% stacked column charts.
- Consider whether the data matches the specified purpose and data requirements 
for the chart type 
- If there is a chart type that allows to use all data, do not select chart types that use only 
part of the data (e.g. do not use column chart, when clustered column chart is also feasible)

Select less than three chart types, if there are less than three appropriate chart types.

2. Is the input data in long format?

3. Determine if the last row of the data contains the sum of all previous rows

    """
    return prompt.strip()


def create_two_column_category_chart_data_selection_prompt(table_headers, chart_message, chart_type,
                                                           header_cell_formats):
    return (
        f""" 
        
        You are provided with a table with the following characteristics: 
        - **Column names:**
        {table_headers}
        
        - ***Each column has a specific format, as described here:**
        {header_cell_formats}
        
        The goal is to create a {chart_type} to support the following message:
        '{chart_message}'
        
        ### Tasks:
        
        Please identify:
        1. Which column should be used as the categories?
        2. Which column should be used as the values? The column names must match exactly the column names that were provided above.
        3. Provide a short descriptive name for the values (e.g., 'Living room size').
        4. Provide a unit for the value (e.g. square meters). If no sensible unit can be found answer with "none". For currency units please always use the ISO currency code e.g. EUR instead of €. 
        5. Does the column selected for the categories have a natural sorting order (for example because it is a time series or categories going from bad to good)?
        
        """
    )


def create_multicolumn_category_chart_data_selection_prompt(table_headers, chart_message, chart_type,
                                                            header_cell_formats):
    return (
        f""" 
        
        You are provided with a table with the following characteristics: the following columns: 
        - **Column names:**
        {table_headers}
        
        - ***Each column has a specific format, as described here:**
        {header_cell_formats}
        
        The goal is to create a {chart_type} to support the following message:
        '{chart_message}'
        
        ### Tasks:
        
        Please identify:
        1. Which column should be used as the categories?
        2. Which columns should be used for the series data? Please list all columns. Do not include columns that contain sums!
        The column names must match exactly the column names that were provided above.
        3. Provide a short descriptive label for the series data (e.g., 'Living room size').
        4. Provide a unit for the series data (e.g. square meters). If no sensible unit can be found answer with "none". For currency units please always use the ISO currency code e.g. EUR instead of €. 
        5. Does the column selected for the categories have a natural sorting order (for example because it is a time series or categories going from bad to good)?
        
        """
    ).strip()


def create_long_format_multicolumn_category_chart_data_selection_prompt(df, core_message, header_cell_formats):
    # Extract column names and descriptions
    columns = df.columns.tolist()

    # Summarize data overview (e.g., range or unique values for each column)
    data_overview = []
    for col in columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            summary = f"Range: {df[col].min()} to {df[col].max()}"
        else:
            unique_vals = df[col].unique()
            summary = f"Examples: {', '.join(map(str, unique_vals[:3]))}"  # Show up to 3 examples
        data_overview.append(f"- {col}: {summary}")
    data_overview_text = "\n".join(data_overview)

    # Include the first 10 rows of the DataFrame
    first_ten_rows = df.head(10).to_string(index=False)

    # Construct the prompt
    prompt = f"""
    I have a table with the following summary characteristics:

    - **Column names:**
    {columns}

    - **Data overview:**
    {data_overview_text}

    - **First 10 rows of the data:**
    {first_ten_rows}

    - ***Each column has a specific format, as described here:**
    {header_cell_formats}

    I want to create a chart from it that supports the following message:
    "{core_message}"

    ### Tasks:

    The data is currently in long format and needs to be pivoted to transform it into wide format.

    1. Determine columns for pivoting
       - What column should be the **index**? (i.e., the x-axis values of the chart)
       - What column should be the **columns**? (i.e., multiple series in the chart)
       - What column should be the **values**? (i.e., the y-axis values of the chart)

    Briefly explain your reasoning before selecting the columns.

    For the index, columns and values please answer with either a column name or a list of column names. 
    The column names must match exactly the column names that were provided above. Do not add quotation marks or anything else.    

    2. Based on the column selected for **values**, provide a short descriptive title
    3. Provide a unit for the value (e.g. square meters). If no sensible unit can be found answer with "none". For currency units please always use the ISO currency code e.g. EUR instead of €.  
    4. Does the column selected for the index (i.e. the x-axis) have a natural sorting order 
    (for example because it is a time series or categories going from bad to good)?

    """
    return prompt.strip()


def create_bubble_chart_data_selection_prompt(table_headers, chart_message, chart_type, header_cell_formats):
    return (
        f""" 
        
        You are provided with a table with the following characteristics: the following columns: 
        - **Column names:**
        {table_headers}
        
        - ***Each column has a specific format, as described here:**
        {header_cell_formats}
        
        The goal is to create a {chart_type} to support the following message:
        '{chart_message}'
        
        ### Tasks:
        
        Please identify:
        1. Which column should be used as the x-axis?
        2. Which column should be used as the y-axis?
        3. Which column should be used for the labels?
        4. Which column should be used for the bubble size? The column names must match exactly the column names that were provided above.
        5. For x- and y-axis and bubble size please provide a descriptive title if applicable (e.g., 'Living room size in square meters' or 'Vehicle sales in EUR').
        For currency units please always use the ISO currency code e.g. EUR instead of €. For percentages please use the % symbol
        6. Additionally, provide a short descriptive title for the chart

        """
    ).strip()
