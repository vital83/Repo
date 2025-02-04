import math
import pandas
import numpy
from datetime import datetime

from . import load_amplitude_tools

SESSION_START_EVENT = "session_start"
AMPLITUDE_SOURCE = "amplitude"


CUSTOM_EVENTS_FOLDER = "./Custom_Events/"
SESSIONS_FOLDER = "./Sessions/"

MIN_EVENT_ROWS_COUNT_IN_PERCENT = 0 # мин. количество строк с Сustom Event в процентах от общего числа строк в полной выборке событий 
MIN_PARAM_COLUMN_FILLING_PERCENT = 100 # # мин. процент заполненности параметра значениями

# названия колонок соответствуют их названиям во входном df
_col_timestamp = "event_time" 
_col_user_id = "user_id" 
_col_session_id = "session_id" 
_col_session_time = "session_time"

_col_event_name = "event_type"

_col_country = "country"
_col_city = "city"
_col_operating_system_version = "os_version"
_col_mobile_model_name = "device_type"
_col_app_version = "version_name"

# этих колонок нет во входящем df - придется сочинить им значения
_col_source_name = "source_name" # по-умолчанию = 'amplitude'
_col_source_medium = "source_medium" # по-умолчанию = 'None'
_col_source_source = "source_source" # по-умолчанию = 'None'

# эту колонку будем добавлять в df - она склеивается из _col_user_id и _col_timestamp обрезанной
# можно попробовать использовать колонку uuid - она уникальная 
_col_unique_session_id = "unique_session_id"

# fixed - может потому что во всех формируемых csv выгрузках он одинаковый
FIXED_FIELDS = [_col_event_name, _col_timestamp, _col_user_id, _col_unique_session_id]

geo_custom_fields = [
    _col_country,
    _col_city
]

device_custom_fields = [
    _col_operating_system_version,
    _col_mobile_model_name
]

source_custom_fields = [
    _col_source_name,
    _col_source_medium,
    _col_source_source
]

app_custom_fields = [
    _col_app_version
]

# названия колонок в нужной последовательности, чтобы записать в правильном формате в выходной файл

OUT_FIXED_CSV_FIELDS = ["event_name", "event_timestamp", "user_pseudo_id", "unique_session_id"]
OUT_SESSIONS_CSV_FIELDS = OUT_FIXED_CSV_FIELDS + ["session_time",
 "country", "city",
  "operating_system_version", "mobile_model_name",
  "source_name", "source_medium", "source_source",
   "app_version"]


def get_input_df(folder, csv_filename):
    df = pandas.read_csv(folder + csv_filename, sep=",")
    return df


def get_user_session_times(user_events, user_id):
    
    if type(user_events) is not pandas.DataFrame:
        return None

    user_sessions = user_events[_col_session_id].unique()
    user_session_times = pandas.DataFrame(columns=[_col_user_id, _col_session_id, _col_session_time])

    for session_id in user_sessions:

        if pandas.isna(session_id):
            continue

        session_events = user_events[user_events[_col_session_id] == session_id].copy()
        session_events.sort_values(_col_timestamp, inplace = True)

        if not is_valid_events_chain(session_events):
            continue

        # session_first_timestamp = int(session_id)
        # session_last_timestamp = int( session_events.tail(1).iloc[0][_col_timestamp] / 1000000)
        session_first_timestamp = int( session_events.head(1).iloc[0][_col_timestamp] / 1000000)
        session_last_timestamp = int(session_events.tail(1).iloc[0][_col_timestamp] / 1000000)
        
        session_time = session_last_timestamp - session_first_timestamp
        session_time = round(session_time)
        if session_time < 0:
            session_time = 0

        user_session_time = pandas.DataFrame(
            [[user_id, session_id, session_time]],
            columns=[_col_user_id, _col_session_id, _col_session_time])
        
        user_session_times = pandas.concat([user_session_times, user_session_time], ignore_index=True, axis=0)

    return user_session_times

def is_valid_events_chain(events):
    TIME_DIFFERENCE_30_MINS = 1800
    events["timestamp_previous"] = events[_col_timestamp].shift()
    events["time_difference"] = (events[_col_timestamp] - events["timestamp_previous"]) / 1000000
    abnormal_events = events[events["time_difference"] > TIME_DIFFERENCE_30_MINS]

    if len(abnormal_events) > 0:
        return False
    
    return True
def convert_datetime_to_nanosecond_column(df, column_name_):
    # df[column_name_].apply(lambda x: str(x.value/100))
    df[column_name_] = (df[column_name_].astype(numpy.int64) / int(1e3)).round().astype(numpy.int64)
    return df 

def convert_string_to_datetime_column(df, column_name_):
    df[column_name_] = pandas.to_datetime(df[column_name_], format='%Y-%m-%d %H:%M:%S.%f')
    return df

def convert_string_datetime_to_nanosecond_column(df, column_name_):
    df[column_name_] = (pandas.to_datetime(df[column_name_], format='%Y-%m-%d %H:%M:%S.%f').astype(numpy.int64) / int(1e3)).round().astype(numpy.int64)
    return df

# wiseman code
START = pandas.Timestamp.now()

#context: OpExecutionContext, 
def perf_interval(label):
    diff = pandas.Timestamp.now() - START
    # print(label, str(diff))
    # context.log.info(label)
    

def simplify_device_fields(events_dataframe):
    events_dataframe[_col_operating_system_version] = events_dataframe.apply(lambda row: simplify_os(row), axis=1)
    return events_dataframe

def simplify_os(row):
    os = str(row[_col_operating_system_version])
    simple_os = os.split(".")[0]
    return simple_os

def remove_rows_with_empty_values(events_dataframe, _col_id):
    events_dataframe[_col_id].replace("", numpy.nan, inplace = True)
    events_dataframe.dropna(subset=[_col_id], inplace = True)
    return events_dataframe

def generate_unique_session_id(row):
    if (row[_col_session_id] is None):
        unique_session_id = str(row[_col_user_id]) + "_" + timestamp_to_session_id(str(row[_col_timestamp]))
    else:
        unique_session_id = str(row[_col_user_id]) + "_" + str(row[_col_session_id])

    return unique_session_id

def timestamp_to_session_id(timestamp_string):
    return timestamp_string[:10]

# end wiseman code

def preprocess_amplitude_dataframe(DATAFRAME):    
    DATAFRAME = convert_string_datetime_to_nanosecond_column(DATAFRAME, _col_timestamp)
    DATAFRAME = DATAFRAME[ (DATAFRAME[_col_user_id] != '') ]
    DATAFRAME = DATAFRAME[ (DATAFRAME[_col_session_id] != -1) ]
    DATAFRAME[_col_unique_session_id] = DATAFRAME.apply(lambda row: generate_unique_session_id(row), axis=1)
    

    # sort to prepare removing duplicate _col_user_id, _col_session_id and keep min value of _col_timestamp
    DATAFRAME.sort_values(by=[_col_user_id, _col_session_id, _col_timestamp])

    return DATAFRAME



def process_sessions(DATAFRAME):
    file_name = SESSION_START_EVENT
    _path_sessions = SESSIONS_FOLDER + file_name + ".csv"
    events = DATAFRAME.copy()
  
    events = simplify_device_fields(events)

    clean_events = events.copy()
    clean_events.set_index(_col_user_id, inplace=True)
    unique_user_ids = clean_events.index.drop_duplicates()

    all_session_times = pandas.DataFrame(columns=[_col_user_id, _col_session_id, _col_session_time])

    perf_interval("Before users cycle ---")

    for user_id in unique_user_ids:
        user_session_times = get_user_session_times(clean_events.loc[user_id], user_id)

        if user_session_times is None:
            continue

        all_session_times = pandas.concat([all_session_times, user_session_times], ignore_index=True, axis=0)
        
    perf_interval("After users cycle ---")

    events.drop_duplicates(subset=[_col_user_id, _col_session_id], inplace=True)
    events = remove_rows_with_empty_values(events, _col_session_id)

    events = pandas.merge(
        events,
        all_session_times,
        how = 'left',
        on = [_col_user_id, _col_session_id])

    events[_col_session_time] = events[_col_session_time].fillna(0)
    events[_col_event_name] = SESSION_START_EVENT
    events[_col_source_name] = AMPLITUDE_SOURCE
    events[_col_source_medium] = None
    events[_col_source_source] = None

    # осталось собрать поля в правильном порядке, переименовать и записать
    events = events[FIXED_FIELDS + [_col_session_time]
                    + geo_custom_fields + device_custom_fields 
                    + source_custom_fields + app_custom_fields].copy()
    
    events.columns = OUT_SESSIONS_CSV_FIELDS
    load_amplitude_tools.make_dir_if_not_exists(SESSIONS_FOLDER)
    if not load_amplitude_tools.clear_folder(SESSIONS_FOLDER): return False
    events.to_csv(_path_sessions, index=False)

    perf_interval("Sessions: Saved " + str(len(events)) + " rows")
    
    return True


def get_column_fill_value_percrentage(DATAFRAME, param_name ):
    column_fill_value_percrentage = 0
    rows_count = DATAFRAME.shape[0]
    DATAFRAME[param_name].replace("", numpy.nan, inplace = True)
    DATAFRAME[param_name].replace("EMPTY", numpy.nan, inplace = True)

    if(rows_count > 0):
        column_fill_value_percrentage = round(100 * DATAFRAME[param_name].notna().sum() / rows_count)

    return column_fill_value_percrentage

def check_param_is_not_empty(events, col_name):
    if '.' in col_name \
        and get_column_fill_value_percrentage(events, col_name) >= MIN_PARAM_COLUMN_FILLING_PERCENT:
        return True    
    return False

def get_event_param_list(DATAFRAME, event_name):
    param_list = []
    events = DATAFRAME[ DATAFRAME[_col_event_name] == event_name ]
    for param in events.columns:
        if check_param_is_not_empty(events, param):
            param_list.append(param)
        
    return param_list

def get_all_unique_custom_events(DATAFRAME):
    clean_events = DATAFRAME.copy()
    clean_events.set_index(_col_event_name, inplace=True)
    unique_event_names = clean_events.index.drop_duplicates()
    
    return unique_event_names


def get_custom_events(DATAFRAME):
    clean_events = DATAFRAME.copy()
    
    min_count = math.ceil(DATAFRAME.shape[0] * MIN_EVENT_ROWS_COUNT_IN_PERCENT / 100)
    unique_event_names_and_counts = clean_events[_col_event_name]               \
                                        .groupby(clean_events[_col_event_name]) \
                                            .filter(lambda x: len(x) >= min_count)       \
                                                .value_counts()
    return unique_event_names_and_counts.index

def process_event(event_name, DATAFRAME):
    EXCLUDED_EVENTS = []
    FIXED_EXCLUDED_PARAMS = ["$insert_id", "$insert_key", "$schema", "adid", "amplitude_attribution_ids", "amplitude_event_type", "amplitude_id", "app", "client_event_time", "client_upload_time", "data_type", "device_brand", "device_carrier", "device_family", "device_id", "device_manufacturer", "device_model", "dma", "event_id", "global_user_properties", "idfa", "ip_address", "is_attribution_event", "language", "library", "location_lat", "location_lng", "os_name", "partner_id", "paying", "platform", "processed_time", "region", "sample_rate", "server_received_time", "server_upload_time", "source_id", "start_version", "user_creation_time", "uuid"]
    EVENT_PARAMS = get_event_param_list(DATAFRAME, event_name)

    print("Processing custom event: " + event_name)
    
    event_data = DATAFRAME[ DATAFRAME[_col_event_name] == event_name ].copy()
    event_data = event_data[FIXED_FIELDS + EVENT_PARAMS].copy()
    event_data.columns = OUT_FIXED_CSV_FIELDS + EVENT_PARAMS
    
    # _path_event_data = path_prefix + _path_data_processed + _custom_events_folder + event_name + "_" + extract_file_marker(file_name) + ".csv"
    _path_event_data = CUSTOM_EVENTS_FOLDER + event_name + ".csv"
    event_data.to_csv(_path_event_data, index=False)
    
    return True

def process_custom_events(DATAFRAME):
        
    unique_events = get_custom_events(DATAFRAME)
    # unique_events = get_all_unique_custom_events(DATAFRAME)

    load_amplitude_tools.make_dir_if_not_exists(CUSTOM_EVENTS_FOLDER)
    if not load_amplitude_tools.clear_folder(CUSTOM_EVENTS_FOLDER): return False    

    for event_config in unique_events:
        process_event(event_config, DATAFRAME)

    return True
