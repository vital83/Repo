import pandas
from dagster import asset

from . import load_amplitude_assets, fetching_tools

@asset
def preprocessed_amplitude_dataframe(amplitude_data_frame) -> pandas.DataFrame:
    DATAFRAME = amplitude_data_frame
    return fetching_tools.preprocess_amplitude_dataframe(DATAFRAME)


# @asset(deps=[preprocessed_amplitude_dataframe])
@asset
def sessions(preprocessed_amplitude_dataframe):
    DATAFRAME = preprocessed_amplitude_dataframe
    fetching_tools.process_sessions(DATAFRAME)

# @asset(deps=[preprocessed_amplitude_dataframe])
@asset
def custom_events(preprocessed_amplitude_dataframe):
    DATAFRAME = preprocessed_amplitude_dataframe
    fetching_tools.process_custom_events(DATAFRAME)

