# 📂 main.py
from integration_ruru_excel826 import flag_close_run_process

if __name__ == "__main__":
    base_month_str = 'M2509'

    # fundType = "public"
    fundType = "private"
    # folder_path = "D:/CommentCheck/702/88/202506基準_連携/04_レポートデータ"
    folder_path = "D:/CommentCheck/702/6yue_reportdata/レポートデータ"

    # ENV = "https://namcheckweb.azurewebsites.net"
    # ENV = "https://namcheckwebrpa.azurewebsites.net"
    ENV = "https://namcheckwebrpa-uat.azurewebsites.net"

    # integration_ruru_excel826.py calling file

    flag_close_run_process(base_month_str, fundType, folder_path, ENV)
