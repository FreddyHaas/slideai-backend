from pydantic import BaseModel


class BarChartDataStructure(BaseModel):
    category: str
    subcategory: str
    value: str
    title: str


class SelectedChartType(BaseModel):
    chartType: str
