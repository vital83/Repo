import zipfile, json, fnmatch, os, shutil, uuid
import pandas, requests, base64, gzip, datetime
from pathlib import Path

EXTRACT_DIR = './temp/'
FILE_PATTERN = '*.json.gz'
OUT_DIR = './out/'
OUT_FILE_NAME_PREFIX = 'amplitude_events_'

def get_datetime_days_ago(days_past):
    today_date = datetime.date.today()
    delta = datetime.timedelta(days = days_past)
    return today_date - delta


def get_amplitude_zip_file(start_date_YYYYMMDD_string, end_date_YYYYMMDD_string, userpass, amplitude_zip_file_name):
    
    # Адрес api метода для запроса get 
    url_param = "https://amplitude.com/api/2/export"
    api_params = {
    "start": start_date_YYYYMMDD_string + "T00",
    "end": end_date_YYYYMMDD_string + "T23"
    }
    
    encoded_userpass = base64.b64encode(userpass.encode()).decode()

    header_params = {
        'GET': '/api/2/export HTTP/1.1',
        'Host': 'amplitude.com',
        "Authorization": "Basic %s" % encoded_userpass
        }
    # Отправляем get request (запрос GET)
    response = requests.get(
        url_param,
        params=api_params,
        headers=header_params
    )

    result_status = response.status_code
    if(result_status == 200):
        with open(amplitude_zip_file_name, "wb") as zip_response:
            zip_response.write(response.content)

    # по факту файл записывается не в хранилище, а в корень проекта
    # это zip архив, внутри которого папка и в ней много файлов с именем вида 374182_2024-03-02_0#0.json.gz
    return amplitude_zip_file_name

def make_data_frame_from_amplitude_zip_file(amplitude_zip_file_name):
    # есть 2 варианта:
    # 1 - использовать локальную файловую систему - пока так и делаем
    # 2 - разобраться как пользоваться временным файловым хранилищем dagster
    # пока много кода на обработку файлов завязано и это не очень удобно, нужно напрямую работать с данными которые в файлах хранятся
    # если переписать - перейти сразу на pandas.data_frame - то будет попроще наверное
    # первая часть - там где скачан архив с небольшими файлами, может его сразу в память распаковывать и добавлять в общий df?
    # сейчас там файлы читаются и добавляются в строку untitled_raw_json_data
    # и потом результат уже передават дальше не как файл, а как df

    file_list = unzip_response(amplitude_zip_file_name)
    
    untitled_raw_json_data = ""
    
    for gz_file_name in file_list:
        with gzip.open(EXTRACT_DIR + gz_file_name, 'rt',encoding='utf-8') as json_file:
                untitled_raw_json_data += json_file.read()
    # файлы были распакованы в каталог c:\Repo\Wiseman\amplitude_to_raw_local\temp\

    json_data = parse_newline_delimited_json(untitled_raw_json_data)
    data_frame = pandas.json_normalize(json_data)
    csv_file_name = get_csv_filename(file_list, OUT_FILE_NAME_PREFIX)
    # make_dir_if_not_exists(OUT_DIR)
    # if not clear_folder(OUT_DIR): return False
    #data_frame.to_csv(OUT_DIR + csv_file_name, index=False, encoding='utf-8')

    '''
    metadata = {
        "num_records": len(data_frame),
        "preview": MetadataValue.md(data_frame.head(10).to_markdown()),
    }
    '''

    # context.add_output_metadata(metadata=metadata)
    return data_frame
    # return OUT_DIR + csv_file_name


def unzip_response(filename_):
    make_dir_if_not_exists(EXTRACT_DIR)
    if not clear_folder(EXTRACT_DIR): return False
    with zipfile.ZipFile(filename_) as response_zip:
         for file in response_zip.infolist():
            if fnmatch.fnmatch(file.filename, FILE_PATTERN):
                response_zip.extract(file.filename, EXTRACT_DIR)
    
    return response_zip.namelist()

def make_dir_if_not_exists(dir_name_):
    if not os.path.exists(dir_name_):
        os.makedirs(dir_name_)
    
    return

def clear_folder(folder_):
    for filename in os.listdir(folder_):
        file_path = os.path.join(folder_, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
            return False
    return True 

def parse_newline_delimited_json(not_valid_js_):
    return json.loads("[" + not_valid_js_.replace("\n", ",")[:-1] + "]")

def get_csv_filename(file_list_, prefix_):
    csv_file_name = str(uuid.uuid4())
    if(len(file_list_) > 0): 
            folder_name = Path(file_list_[0]).parent.name
            file_name = Path(file_list_[0]).stem
            csv_file_name = file_name.split('_')[1].replace('-', '')
    
    return prefix_ + csv_file_name + '.csv'
