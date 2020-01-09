import datetime
import time
import os
import pyexcel as pe
import traceback

def _get_appdata_path():
    import ctypes
    from ctypes import wintypes, windll
    CSIDL_APPDATA = 26
    _SHGetFolderPath = windll.shell32.SHGetFolderPathW
    _SHGetFolderPath.argtypes = [wintypes.HWND,
                                 ctypes.c_int,
                                 wintypes.HANDLE,
                                 wintypes.DWORD,
                                 wintypes.LPCWSTR]
    path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    result = _SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, path_buf)
    return path_buf.value.replace('Roaming', 'Local')

def dropbox_home():
    from platform import system
    import base64
    import os.path
    _system = system()
    if _system in ('Windows', 'cli'):
        host_db_path = os.path.join(_get_appdata_path(),
                                    'Dropbox',
                                    'host.db')
    elif _system in ('Linux', 'Darwin'):
        host_db_path = os.path.expanduser('~'
                                          '/.dropbox'
                                          '/host.db')
    else:
        raise RuntimeError('Unknown system={}'
                           .format(_system))
    if not os.path.exists(host_db_path):
        raise RuntimeError("Config path={} doesn't exists"
                           .format(host_db_path))
    with open(host_db_path, 'r') as f:
        data = f.read().split()

    return base64.b64decode(data[1]).decode()


def get_cards():
    from trello import TrelloClient
    client = TrelloClient(
    api_key='2b65b16a238feaf9885892c35fb586dc',
    api_secret='312165be7d3ccd4174d9bba9686300aed81b5633bb8aebce4cc844073a8237bb',
    token='ec76037b89fa2fc6beecbd9e979e8072541819ea1e46bd247ce8cac1d0cb5046',
    token_secret=''
    )
    card_list = []
    print(dir(client.http_service.status_codes))
    try:
        all_boards = client.list_boards()
        mcl_board = all_boards[2]
        done_list = mcl_board.list_lists()[7]
        for card in done_list.list_cards():
            if "BB - " in card.name:
                if '[' not in card.name:
                    exit(card.name + ' has no point value.')
                card_list.append([datetime.datetime.strftime(card.latestCardMove_date, '%y-%m-%d'), str(card.name).split('[')[0], str(card.name).split('[')[1]])
        return sorted(card_list)
    except Exception as e:
            print(traceback.print_exc())
            exit(e)


def add_rows_to_excel(rated_list):
    print('1')
    dropbox_dir = str(dropbox_home()) + '\\SZ Team\\Brian\\Velocity\\'
    print('2')
    todays_date = datetime.datetime.strftime(datetime.datetime.now(), '%y-%m-%d')
    print('3')
    try:
        print('4')
        sheet = pe.get_sheet(file_name=dropbox_dir + 'brian_velocity.xlsx')
        #print(rated_list)
        for task in rated_list:
            print('6')
            sheet.row += task
        sheet.save_as(dropbox_dir + 'archive\\' + todays_date + '- brian_velocity.xlsx')
        sheet.save_as(dropbox_dir + 'brian_velocity.xlsx')

    except Exception as e:
        print(traceback.print_exc())
        exit(e)

        

    

# Start of program
card_list = get_cards()
print(card_list)
add_rows_to_excel(card_list)
 