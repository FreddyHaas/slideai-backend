from enum import Enum
from typing import List

from pydantic import BaseModel

from enum import Enum


# ToDo Charttypen auskommentiert
class ChartType(Enum):
    COLUMN = (
        "column_chart",
        True,
        "Compare categories or show trends over time",
        "Exactly two columns as data input required, not suited for large number of data points such as time series data"
    )
    COLUMN_CLUSTERED = (
        "clustered_column_chart",
        False,
        "Compare multiple categories side by side for each group",
        "Multiple columns as data input required, the groups and categories should be non numeric"
    )
    COLUMN_STACKED = (
        "stacked_column_chart",
        False,
        "Compare multiple categories side by side for each group and focus on the cumulative value of the categories",
        "Multiple columns as data input required, the groups and categories should be non numeric"
    )
    COLUMN_STACKED_100 = (
        "100_percent_stacked_column_chart",
        False,
        "Compare multiple categories side by side for each group and focus on proportions of subcategories as a percentage of the total within categories",
        "Multiple columns as data input required, the groups and categories should be non numeric"
    )
    PIE = (
        "pie_chart",
        True,
        "Show proportions of a whole",
        "Exactly two columns as data input required, columns should have no more than 5 entries, otherwise consider column chart"
    )
    AREA_STACKED_100 = (
        "100_percent_stacked_area_chart",
        False,
        "Display proportions of subcategories over time as percentages",
        "Number of values on the x-axis (i.e. categories) must be ordinal and should be more than 10 otherwise consider stacked column chart, works well with time series data"
    )
    LINE = (
        "line_chart",
        True,
        "Display trends over time or continuous data",
        "Number of values on the x-axis (i.e. categories) must be ordinal and should be more than 10 otherwise consider column chart, works well with time series data"
    )
    BUBBLE = (
        "bubble_chart",
        False,
        "Show relationships between three numeric variables",
        "Three columns as input required (X, Y, and bubble size), all three values must be numeric"
    )

    def __init__(self, value, works_with_two_input_columns, purpose, data_input):
        self._value_ = value
        self.works_with_two_input_columns = works_with_two_input_columns
        self.purpose = purpose
        self.data_input = data_input

    @classmethod
    def get_all_chart_types_with_two_columns_input(cls):
        return [member for member in cls if member.works_with_two_input_columns]

    @classmethod
    def get_all(cls):
        return list(cls)


class BarOrColumnDataStructure(BaseModel):
    category: str
    value: str
    title: str


class LineOrClusteredColumnChartDataStructure(BaseModel):
    category: str
    series: List[str]
    title: str


class ClusteredBarOrColumnDataStructure(BaseModel):
    category: str
    subcategory: str
    value: str
    title: str


class BubbleChartDataStructure(BaseModel):
    labels_column: str
    x_axis_column: str
    x_axis_title: str
    x_axis_is_percentage: bool
    y_axis_column: str
    y_axis_title: str
    y_axis_is_percentage: bool
    bubble_size_column: str
    bubble_size_title: str
    title: str


class PieChartDataStructure(BaseModel):
    category_column: str
    value_column: str
    title: str


class SelectedChartType(BaseModel):
    reasonForSelectedChartType: str
    chartType: str
    lastLineIncludesSum: bool


class TablePivot(BaseModel):
    data_structure_analysis: str
    long_format: bool
    explain_pivot: str
    index: str
    columns: str
    values: str
