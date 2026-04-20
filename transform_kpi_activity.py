import pandas as pd
import gspread
import warnings ; import os; import time ; from datetime import datetime; import sys
warnings.filterwarnings('ignore')


def transform_kpi(path):
    df = pd.read_excel(path)
    rename_cols = {
    'Điểm ghi nhận theo tiêu chí' : 'Ghé thăm KH',
    'Unnamed: 10' : 'Ghé thăm TN',
    'Unnamed: 11' : 'Tạo mới TN',
    'Unnamed: 12' : 'KH mới',
    'Unnamed: 13' : 'Đơn hàng mới'
}
    df = df.iloc[1:,:]
    df.rename(columns=rename_cols, inplace=True)
    COL_TO_CHANGE_TYPE = ['Ghé thăm KH', 'Ghé thăm TN', 'Tạo mới TN', 'KH mới', 'Đơn hàng mới']
    for col in COL_TO_CHANGE_TYPE:
        df[col] = df[col].astype(int)
    df = df.fillna("")
    return df

def transform_work(path):
    df = pd.read_excel(path)
    col_names = ['Tên Nhân Viên', 'Phòng Ban', 'Số Công', 'Số Giờ', 'Nghỉ', 'Giải Trình', 'Đi Muộn', 'Về Sớm', 'Số Ngày Chấm Công' ]
    df_test = df.iloc[3:, [1,2,3,4, 35,36,37,38,39]]
    df_test.columns = col_names
    col_change_types = ['Số Công', 'Số Giờ', 'Giải Trình', 'Đi Muộn', 'Về Sớm', 'Số Ngày Chấm Công']
    for col in col_change_types:
        df_test[col] = df_test[col].astype(float).astype(int)
    df_test['Tổng đi muộn về sớm'] = df_test['Đi Muộn'].astype(int) + df_test['Về Sớm'].astype(int)
    df_test.fillna("")
    return df_test

def push_data(df ,sheet_id, sheet_name):
    path_to_json =os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
    gc = gspread.service_account(filename=path_to_json) 
    sh = gc.open_by_key(sheet_id)
    worksheet = sh.worksheet(sheet_name)
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

if __name__ == '__main__':
    # list base dir
    try :
        if getattr(sys, 'frozen', False):
        # Nếu là file .exe, lấy thư mục chứa file .exe đó
            base_dir = os.path.dirname(sys.executable)
        else:
        # Nếu là file .py bình thường, lấy thư mục chứa file .py
            base_dir = os.path.dirname(os.path.abspath(__file__))
        path_json = os.path.join(base_dir, 'sheet_push_data.json')
        
        print("Đang khởi tạo môi trường ...")
        # Initialize env 
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = path_json

        # file paths for extract
        transform_kpi_path = os.path.join(base_dir, 'Báo cáo thực hiện - KPI hoạt động.xlsx')
        transform_work_path = os.path.join(base_dir, 'Báo cáo chấm công.xlsx')


        # sheet id for push
        sheet_id = '1Wr74ltOP0lDHUXdJVieWSlxq2-aUAGJJS2LMkZqrnvk'

        # sheet name for each data
        sheet_name_kpi = 'K P I Activity'
        sheet_name_work = 'W O R K I N G D A Y S'

        # push data KPI
        push_data(transform_kpi(transform_kpi_path), sheet_id=sheet_id, sheet_name=sheet_name_kpi)
        print(f"{datetime.now()} || Đẩy data thành công data từ file {transform_kpi_path} vào sheet {sheet_name_kpi}")

        # push data work hours
        push_data(transform_work(transform_work_path), sheet_id=sheet_id, sheet_name=sheet_name_work)
        print(f"{datetime.now()} || Đẩy data thành công data từ file {transform_work_path} vào sheet {sheet_name_work}")


        print("Thành công !")

        time.sleep(3)
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")
    
    # Dòng này giữ cửa sổ lại
    input("\nNhấn Enter để thoát...")


