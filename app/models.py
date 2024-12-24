from enum import Enum
from typing import List

from pydantic import BaseModel


class ChartType(Enum):
    BAR = "bar_chart"
    BAR_CLUSTERED = "clustered_bar_chart"
    BAR_STACKED = "stacked_bar_chart"
    BAR_STACKED_100 = "100_percent_stacked_bar_chart"
    COLUMN = "column_chart"
    COLUMN_CLUSTERED = "clustered_column_chart"
    COLUMN_STACKED = "stacked_column_chart"
    COLUMN_STACKED_100 = "100_percent_stacked_column_chart"
    PIE = "pie_chart"
    DOUGHNUT = "doughnut_chart"
    AREA_STACKED = "area_chart"
    AREA_STACKED_100 = "100_percent_stacked_area_chart"
    LINE = "line_chart"
    BUBBLE = "bubble_chart"


class BarOrColumnDataStructure(BaseModel):
    category: str
    value: str
    title: str


class LineChartDataStructure(BaseModel):
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
    chartType: str
    lastLineIncludesSum: bool
