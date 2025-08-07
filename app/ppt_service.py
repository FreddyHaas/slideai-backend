import datetime
import os
import subprocess
from typing import Optional

import numpy as np
from pptx import Presentation

from chart_factory import create_clustered_column_chart, create_clustered_bar_chart, create_stacked_column_chart, \
    create_100_percent_stacked_column_chart, create_line_chart, create_column_chart, create_bar_chart, \
    create_pie_chart, create_doughnut_chart, create_bubble_chart, create_stacked_bar_chart
from openai_adapter import _query_openai
from prompt_factory import create_two_column_category_chart_data_selection_prompt, \
    create_multicolumn_category_chart_data_selection_prompt, \
    create_long_format_multicolumn_category_chart_data_selection_prompt, create_chart_selection_prompt, \
    create_bubble_chart_data_selection_prompt
from models import MultiColumnDataStructure, PowerpointCreationResponse, SelectedChartType, ChartType, \
    TwoColumnDataStructure, \
    LongFormatDataStructure, BubbleChartDataStructure, RoundingPrecision

MOCK_AI_API_CALLS = False

current_dir = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.path.join(current_dir, "template.pptx")


# Data transformation
def _normalize_values_to_percentages_multi_columns(dataframe, series: list[str]):
    percentage_dataframe = dataframe.copy()
    percentage_dataframe[series] = dataframe[series].div(dataframe[series].sum(axis=1),
                                                         axis=0) * 100  # Convert to percentage
    return percentage_dataframe


def _normalize_values_to_percentages_single_column(dataframe, value: str):
    total = dataframe[value].sum()
    percentage_dataframe = dataframe.copy()
    percentage_dataframe[value] = (dataframe[value] / total)

    return percentage_dataframe


# Data ingestion

def _sort_descending(two_column_dataframe, two_column_chart_information):
    return two_column_dataframe.sort_values(by=two_column_chart_information.value,
                                            ascending=False) if not two_column_chart_information.has_natural_sorting_order else two_column_dataframe


def _determine_rounding_precision(df, columns) -> RoundingPrecision:
    order_of_magnitude = 0
    decimal_place = 0

    for column in columns:
        # Calculate the median and determine the order of magnitude
        median = df[column].median()
        if median == 0:  # Avoid log10 issues with zero
            print(f"Median of column '{column}' is zero. Skipping.")
            continue

        order_of_magnitude_column = int(np.floor(np.log10(abs(median))))
        if order_of_magnitude_column > order_of_magnitude:
            order_of_magnitude = order_of_magnitude_column

    if order_of_magnitude in [0, 1, 3, 4, 6, 7, 9, 10]:
        divisor = 1
        # 1.000 / 10.000
        if order_of_magnitude in [3, 4]:
            divisor = 1000
        # 1.000.000 / 10.000.000
        if order_of_magnitude in [6, 7]:
            divisor = 1000000
        # 1.0000.000.000 / 10.000.000.000
        if order_of_magnitude in [9, 10]:
            divisor = 1000000000
        for column in columns:
            divisibility = (df[column] / divisor) % 1 == 0
            if not divisibility.all():
                decimal_place = 1

    return RoundingPrecision(
        order_of_magnitude=order_of_magnitude,
        decimal_place=decimal_place
    )


def _convert_pptx_to_pdf(pptx_file):

    command = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        pptx_file
    ]

    subprocess.run(command, check=True)


# Main function
def create_chart(df, header_cell_formats: dict, chart_core_message: str, uuid):
    selected_two_column_charts = ChartType.get_two_column_charts()
    selected_multi_column_charts = ChartType.get_multi_column_charts()
    all_charts = ChartType.get_all()

    presentation = Presentation(TEMPLATE_PATH)

    df_headers = df.columns.tolist()
    has_more_than_two_headers = len(df_headers) > 2

    # Select chart
    chart_selection_prompt = create_chart_selection_prompt(
        df=df,
        chart_options=all_charts if has_more_than_two_headers else selected_two_column_charts,
        core_message=chart_core_message,
        header_cell_formats=header_cell_formats)

    selected_chart_type = _query_openai(message=chart_selection_prompt, response_model=SelectedChartType)
    # MOCK
    # selected_chart_type = SelectedChartType(
    #     reason_for_selected_chart_types="some reason",
    #     chart_types=[ChartType.COLUMN.value],
    #     is_in_long_format=False,
    #     last_line_includes_sum=False
    # )

    print(selected_chart_type.reason_for_selected_chart_types)
    selected_charts = selected_chart_type.chart_types
    is_long_format = selected_chart_type.is_in_long_format

    # Prepare data
    df.columns = df.columns.astype(str)

    if selected_chart_type.last_line_includes_sum:
        df = df.drop(df.index[-1])

    category_charts = ChartType.get_category_chart_names()
    multi_category_charts = ChartType.get_multi_category_chart_names()

    selected_two_column_charts = list(set(selected_charts).intersection(category_charts))
    selected_multi_column_charts = list(set(selected_charts).intersection(multi_category_charts))

    multi_column_dataframe = None
    multi_column_chart_information: Optional[MultiColumnDataStructure]
    multi_column_rounding_precision: Optional[RoundingPrecision]

    two_column_dataframe = None
    two_column_chart_information: Optional[TwoColumnDataStructure]
    two_column_rounding_precision: Optional[RoundingPrecision]

    bubble_dataframe = None
    bubble_chart_information: Optional[BubbleChartDataStructure]

    if selected_multi_column_charts:
        try:
            if is_long_format:
                data_selection_prompt = create_long_format_multicolumn_category_chart_data_selection_prompt(df=df,
                                                                                                            core_message=chart_core_message,
                                                                                                            header_cell_formats=header_cell_formats
                                                                                                            )

                selected_data = _query_openai(
                    message=data_selection_prompt,
                    response_model=LongFormatDataStructure
                )

                multi_column_dataframe = df.pivot(
                    index=selected_data.index,
                    columns=selected_data.columns,
                    values=selected_data.values
                )

                multi_column_dataframe = multi_column_dataframe.reset_index()
                multi_column_dataframe.columns = multi_column_dataframe.columns.astype(str)

                column_headers = multi_column_dataframe.columns.tolist()

                multi_column_chart_information = MultiColumnDataStructure(
                    category=column_headers[0],
                    series=column_headers[1:],
                    axis_label=selected_data.title,
                    axis_unit=selected_data.unit,
                    has_natural_sorting_order=selected_data.has_natural_sorting_order
                )

            else:

                data_selection_prompt = (
                    create_multicolumn_category_chart_data_selection_prompt(
                        df_headers,
                        chart_core_message,
                        "clustered column chart",
                        header_cell_formats
                    )
                )

                multi_column_chart_information = _query_openai(
                    message=data_selection_prompt,
                    response_model=MultiColumnDataStructure
                ) if not MOCK_AI_API_CALLS else MultiColumnDataStructure(
                    category="Year",
                    series=["USA", "China"],
                    title="some title",
                    has_natural_sorting_order=False
                )

                multi_column_dataframe = df.groupby(multi_column_chart_information.category, as_index=False).sum()
                multi_column_dataframe.columns = multi_column_dataframe.columns.astype(str)

            if not multi_column_chart_information.has_natural_sorting_order:
                row_sums = multi_column_dataframe[multi_column_chart_information.series].sum(axis=1)
                multi_column_dataframe = multi_column_dataframe.loc[row_sums.sort_values(ascending=True).index]

            multi_column_rounding_precision = _determine_rounding_precision(
                multi_column_dataframe,
                multi_column_chart_information.series
            )

        except Exception as exception:
            selected_charts = list(set(selected_charts) - set(selected_multi_column_charts))
            print(str(exception))

    if selected_two_column_charts:
        try:
            data_selection_prompt = create_two_column_category_chart_data_selection_prompt(
                table_headers=df_headers,
                chart_message=chart_core_message,
                chart_type="column chart",
                header_cell_formats=header_cell_formats)

            two_column_chart_information = _query_openai(
                message=data_selection_prompt,
                response_model=TwoColumnDataStructure
            ) if not MOCK_AI_API_CALLS else (
                TwoColumnDataStructure(
                    category="Market",
                    value="Units sold",
                    axis_label="Units sold",
                    axis_unit="none",
                    has_natural_sorting_order=False
                )
            )

            two_column_dataframe = df.groupby(two_column_chart_information.category, as_index=False).sum()
            two_column_dataframe.columns = two_column_dataframe.columns.astype(str)

            if not two_column_chart_information.has_natural_sorting_order:
                two_column_dataframe = two_column_dataframe.sort_values(by=two_column_chart_information.value)

            two_column_rounding_precision = _determine_rounding_precision(
                two_column_dataframe,
                [two_column_chart_information.value]
            )

        except Exception as exception:
            selected_charts = list(set(selected_charts) - set(selected_two_column_charts))
            print(str(exception))

    if ChartType.BUBBLE.value in selected_charts:
        try:
            data_selection_prompt = create_bubble_chart_data_selection_prompt(df_headers, chart_core_message,
                                                                              "bubble chart",
                                                                              header_cell_formats)
            bubble_chart_information = _query_openai(
                message=data_selection_prompt,
                response_model=BubbleChartDataStructure
            ) if not MOCK_AI_API_CALLS else (
                BubbleChartDataStructure(
                    labels_column="Market",
                    x_axis_column="Market share",
                    y_axis_column="Market growth",
                    x_axis_is_percentage=True,
                    y_axis_is_percentage=True,
                    x_axis_title="Market share (%)",
                    y_axis_title="Market growth (%)",
                    bubble_size_column="Market size",
                    bubble_size_title="Some title",
                    title="Market size in EUR"
                )
            )

            bubble_dataframe = df[[bubble_chart_information.labels_column,
                                   bubble_chart_information.x_axis_column,
                                   bubble_chart_information.y_axis_column,
                                   bubble_chart_information.bubble_size_column]].copy()

            if bubble_chart_information.x_axis_is_percentage:
                bubble_dataframe[bubble_chart_information.x_axis_column] *= 100

            if bubble_chart_information.y_axis_is_percentage:
                bubble_dataframe[bubble_chart_information.y_axis_column] *= 100

            bubble_dataframe.columns = bubble_dataframe.columns.astype(str)
        except Exception as exception:
            selected_charts = list(set(selected_charts) - set(ChartType.BUBBLE.value))
            print(str(exception))

    for chart in selected_charts:
        match chart:
            # Multi column charts
            case ChartType.COLUMN_CLUSTERED.value:
                create_clustered_column_chart(
                    presentation=presentation,
                    dataframe=multi_column_dataframe,
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=multi_column_rounding_precision
                )
                create_clustered_bar_chart(
                    presentation=presentation,
                    dataframe=multi_column_dataframe,
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=multi_column_rounding_precision
                )
            case ChartType.COLUMN_STACKED.value:
                create_stacked_column_chart(
                    presentation=presentation,
                    dataframe=multi_column_dataframe,
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=multi_column_rounding_precision
                )
                create_stacked_bar_chart(
                    presentation=presentation,
                    dataframe=multi_column_dataframe,
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=multi_column_rounding_precision
                )
            case ChartType.COLUMN_STACKED_100.value:
                create_100_percent_stacked_column_chart(
                    presentation=presentation,
                    dataframe=_normalize_values_to_percentages_multi_columns(multi_column_dataframe,
                                                                             multi_column_chart_information.series),
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message
                )
            case ChartType.LINE.value:
                create_line_chart(
                    presentation=presentation,
                    dataframe=multi_column_dataframe,
                    chart_information=multi_column_chart_information,
                    chart_core_message=chart_core_message
                )
            # Two column charts
            case ChartType.COLUMN.value:
                create_column_chart(
                    presentation=presentation,
                    dataframe=two_column_dataframe,
                    chart_information=two_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=two_column_rounding_precision
                )
                create_bar_chart(
                    presentation=presentation,
                    dataframe=two_column_dataframe,
                    chart_information=two_column_chart_information,
                    chart_core_message=chart_core_message,
                    rounding_precision=two_column_rounding_precision
                )
            case ChartType.PIE.value:

                percentage_dataframe = _normalize_values_to_percentages_single_column(two_column_dataframe,
                                                                                      two_column_chart_information.value)
                sorted_percentage_dataframe = _sort_descending(percentage_dataframe, two_column_chart_information)
                create_pie_chart(
                    presentation=presentation,
                    dataframe=sorted_percentage_dataframe,
                    chart_information=two_column_chart_information,
                    chart_core_message=chart_core_message
                )
                create_doughnut_chart(
                    presentation=presentation,
                    dataframe=sorted_percentage_dataframe,
                    chart_information=two_column_chart_information,
                    chart_core_message=chart_core_message
                )
            case ChartType.BUBBLE.value:
                create_bubble_chart(
                    presentation=presentation,
                    dataframe=bubble_dataframe,
                    chart_information=bubble_chart_information,
                    chart_core_message=chart_core_message
                )
    if len(presentation.slides) < 1:
        raise Exception("Unable to create chart")

    presentation_name = f"{uuid}_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
    ppt_path = f"{presentation_name}.pptx"
    presentation.save(ppt_path)

    _convert_pptx_to_pdf(ppt_path)

    return PowerpointCreationResponse(
        presentation_name=presentation_name,
    )
