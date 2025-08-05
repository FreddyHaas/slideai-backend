from enum import Enum
from typing import List

from pydantic import BaseModel

from enum import Enum


class DataStructure(Enum):
    CATEGORY = "category"
    MULTI_CATEGORY = "multi_category"
    BUBBLE = "bubble"
    XY = "xy"


class ChartType(Enum):
    COLUMN = (
        "column_chart",
        True,
        "Compare categories or show trends over time",
        "At least one numeric column, not suited for large number of data points such as time series data",
        DataStructure.CATEGORY
    )
    COLUMN_CLUSTERED = (
        "clustered_column_chart",
        False,
        "Compare multiple categories side by side for each group",
        "Multiple columns, the groups and categories should be non numeric, the values have to be"
        "numeric and in the same unit",
        DataStructure.MULTI_CATEGORY
    )
    COLUMN_STACKED = (
        "stacked_column_chart",
        False,
        "Compare multiple categories side by side for each group, focus on the cumulative value of the categories",
        "Multiple columns, the groups and categories should be non numeric, the values have to be"
        "numeric and in the same unit",
        DataStructure.MULTI_CATEGORY
    )
    COLUMN_STACKED_100 = (
        "100_percent_stacked_column_chart",
        False,
        "Compare multiple categories side by side for each group and focus on proportions of subcategories as a percentage of the total within categories",
        "Multiple columns, the groups and categories should be non numeric, the values have to be"
        "numeric and in the same unit",
        DataStructure.MULTI_CATEGORY
    )
    PIE = (
        "pie_chart",
        True,
        "Show proportions of a whole",
        "At least one numeric column, columns should have no more than 5 entries, otherwise consider column chart",
        DataStructure.CATEGORY
    )
    # AREA_STACKED_100 = (
    #     "100_percent_stacked_area_chart",
    #     False,
    #     "Display proportions of subcategories over time as percentages",
    #     "Number of values on the x-axis (i.e. categories) must be ordinal and should be more than 10 otherwise consider stacked column chart, works well with time series data",
    #     DataStructure.CATEGORY
    # )
    LINE = (
        "line_chart",
        True,
        "Display trends over time or continuous data",
        "Number of values on the x-axis (i.e. categories) must be ordinal and should be more than 10 otherwise consider column chart, works well with time series data",
        DataStructure.MULTI_CATEGORY
    )
    BUBBLE = (
        "bubble_chart",
        False,
        "Show relationships between three numeric variables",
        "Three columns as input required (X, Y, and bubble size), all three values must be numeric, they can have different units and magnitudes",
        DataStructure.BUBBLE
    )

    def __init__(self, value, works_with_two_input_columns, purpose, data_input, data_structure):
        self._value_ = value
        self.works_with_two_input_columns = works_with_two_input_columns
        self.purpose = purpose
        self.data_input = data_input
        self.data_structure = data_structure

    @classmethod
    def get_two_column_charts(cls):
        return [member for member in cls if member.works_with_two_input_columns]

    @classmethod
    def get_multi_column_charts(cls):
        return [member for member in cls if not member.works_with_two_input_columns]

    @classmethod
    def get_category_chart_names(cls):
        return [member.value for member in cls if member.data_structure == DataStructure.CATEGORY]

    @classmethod
    def get_multi_category_chart_names(cls):
        return [member.value for member in cls if member.data_structure == DataStructure.MULTI_CATEGORY]

    @classmethod
    def get_all(cls):
        return list(cls)


class TwoColumnDataStructure(BaseModel):
    category: str
    value: str
    axis_label: str
    axis_unit: str
    has_natural_sorting_order: bool


class MultiColumnDataStructure(BaseModel):
    category: str
    series: List[str]
    axis_label: str
    axis_unit: str
    has_natural_sorting_order: bool


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


class SelectedChartType(BaseModel):
    reason_for_selected_chart_types: str
    chart_types: List[str]
    is_in_long_format: bool
    last_line_includes_sum: bool


class LongFormatDataStructure(BaseModel):
    explain_column_selection: str
    index: str
    columns: str
    values: str
    title: str
    unit: str
    has_natural_sorting_order: bool


class DataValidationRequest(BaseModel):
    data: str


class DataValidationResponse(BaseModel):
    is_valid: bool
    validation_hints: list[str]


class RoundingPrecision(BaseModel):
    order_of_magnitude: int
    decimal_place: int


class PowerpointCreationResponse(BaseModel):
    presentation_name: str
