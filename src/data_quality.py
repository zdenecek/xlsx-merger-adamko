import pandera as pa
from pandera import Column, Check

# Define the schema
schema = pa.DataFrameSchema({
    "Filename": Column(pa.String),
    "BIRTH_YEAR": Column(pa.Int32, Check.in_range(1900, 2023), nullable=True),
    "BIRTH_MNTH": Column(pa.Int32, Check.in_range(1, 12), nullable=True),
    "BIRTH_DAY": Column(pa.Int32, Check.in_range(1, 31), nullable=True),
    "BIRTH_WGHT": Column(pa.Float64, Check.in_range(300, 6000), nullable=True),
    "SEX": Column(pa.Int32, Check.isin([1, 2]), nullable=True),
    # Add more columns as needed
})

def validate_data(merged_df, schema):
    try:
        validated_df = schema.validate(merged_df, lazy=True)
        return validated_df, None
    except pa.errors.SchemaErrors as err:
        # Collect error messages
        error_df = err.failure_cases
        return None, error_df