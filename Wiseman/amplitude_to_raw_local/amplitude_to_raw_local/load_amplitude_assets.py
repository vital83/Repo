import pandas, requests, base64, gzip, os

from dagster import OpExecutionContext, EnvVar, MetadataValue, asset, op


from . import load_amplitude_tools

@asset
def amplitude_zip_file(context: OpExecutionContext):
    # берем дату 3 дня назад
    
    start_date_YYYYMMDD_string = load_amplitude_tools.get_datetime_days_ago(3).strftime('%Y%m%d')
    end_date_YYYYMMDD_string = load_amplitude_tools.get_datetime_days_ago(3).strftime('%Y%m%d')
       
    userpass = EnvVar("AMPLITUDE_HUMANIFY_USERPASS")
    # hostname = os.getenv("COMPUTERNAME")
    # context.log.info(f"Hostname is {hostname}")
 
    amplitude_zip_file_name = "response.zip"
    
    return load_amplitude_tools.get_amplitude_zip_file(start_date_YYYYMMDD_string,
                                                        end_date_YYYYMMDD_string,
                                                         userpass,
                                                          amplitude_zip_file_name)



# context: AssetExecutionContext
@asset
def amplitude_data_frame(amplitude_zip_file) -> pandas.DataFrame:
     amplitude_zip_file_name = amplitude_zip_file
     return load_amplitude_tools.make_data_frame_from_amplitude_zip_file(amplitude_zip_file_name)

'''
@asset
def amplitudde_preprocessed_dataframe(deps=[amplitude_csv_file]):
    return

@op
def fetch_session_start():
    return

@op
def fetch_custom_events():
    return

@job
def run_fetcher(deps=[amplitudde_preprocessed_dataframe]):
    fetch_session_start(amplitudde_preprocessed_dataframe)
    fetch_custom_events(amplitudde_preprocessed_dataframe)
'''
