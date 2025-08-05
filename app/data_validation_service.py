from app.models import DataValidationResponse


def fun_validate(df) -> DataValidationResponse:
    is_valid = True
    validation_hints = []

    df.columns = df.columns.astype(str)
    headers = df.columns.tolist()

    # Check for null or empty headers
    if any("_EMPTY" in header or header.strip() == "" for header in headers):
        is_valid = False
        validation_hints.append("Data contains empty headers. Please provide valid headers.")

    # Check if the first row contains unique headers
    if len(headers) != len(set(headers)) and is_valid is True:
        is_valid = False
        validation_hints.append("Excel contains duplicated headers. Please ensure your headers are unique")

    # Check for missing values and locate them
    missing_data = df.isnull()
    if missing_data.values.any():
        is_valid = False
        missing_locations = missing_data.stack()[missing_data.stack()].index.tolist()
        missing_values_hint = "Excel contains missing values. Please fill missing values at: "
        for row, col in missing_locations:
            missing_values_hint = missing_values_hint + f"[row number: {row + 2}, column header: {col}]"
        validation_hints.append(missing_values_hint)

    # Check consistent formatting
    inconsistent_columns = []
    number_columns = False
    for column in df.columns:
        # Get the set of types for this column
        column_types = set(df[column].dropna().apply(type))

        if any(issubclass(t, (int, float)) for t in column_types):
            number_columns = True

        # If there are more than one type, it's inconsistent
        if len(column_types) > 1:
            inconsistent_columns.append(column)

    if inconsistent_columns:
        is_valid = False
        validation_hints.append(
            f"The following columns are formatted inconsistently (e.g. contain text and numbers): {", ".join(inconsistent_columns)}")

    if not number_columns:
        is_valid = False
        validation_hints.append(
            f"Could not find any numbers in your data - please check your formatting (e.g. remove Units from entries)"
        )

    return DataValidationResponse(
        is_valid=is_valid,
        validation_hints=validation_hints
    )
