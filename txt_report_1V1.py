import os.path
import zipfile

import pandas as pd
import xlwings as xw
from datetime import datetime
from txt_print import txt_print


# template_path = r'D:\Python\pythonProject\venv\AAC_Yield_Summary_Report\X2061 DVT DOE+Main Configs Yield Summary_20210701.xlsx'
# txt_path = r'D:\Python\pythonProject\venv\AAC_Yield_Summary_Report\txt'

def txt_report(template_path, txt_path, text_info):
    txt_list = []
    for file in os.listdir(txt_path):
        if os.path.splitext(file)[-1] == '.txt':
            txt_list.append(file)

    print_info = f'Start -> Total {len(txt_list)} TXT Files in the Path'
    txt_print(text_info, 'tag1', print_info, 50, 'Cyan1', 'Times', 10)

    config_list = []
    station_list = []
    overlay_list = []
    for i in txt_list:
        k = i.split('_')
        config_list.append(k[-1].split('.')[0])
        station_list.append(k[2])
        overlay_list.append(k[-4])

    file_matrix = pd.DataFrame([config_list, station_list, overlay_list],
                               columns=txt_list)  # 第一行为Config，第二行为station，第三行为Overlay
    file_matrix.sort_values(by=[0], axis=1, inplace=True)

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(template_path)
    ws = wb.sheets[1]

    dt = datetime.now()
    now_date = dt.strftime('%m-%d')
    now_date_2 = dt.strftime('%Y%m%d')

    start_col_num = 3

    while ws[1, start_col_num].value != None:
        start_col_num = start_col_num + 1

    start_col_num = start_col_num - 1
    # file_last = pd.Series(['', '', ''])

    config_flag = 1
    config_file_list = []
    for i in file_matrix.columns:
        config_file_list.append(i)
        index_series_data = pd.Series(ws.range('B1:B400').value)
        index_series = index_series_data[index_series_data.values == file_matrix[i][
            1]].index.tolist()  # station位置index，第一个是良率表格，第二个是Result，第三个是Retest
        if config_flag:
            start_col_num = start_col_num + 1
            ws.api.Columns(start_col_num + 1).Insert()
            ws[1, start_col_num].value = file_matrix[i][0]
            ws[2, start_col_num].value = now_date
            ws[3, start_col_num].value = file_matrix[i][0]
            ws[1:3, start_col_num].color = [255, 239, 213]
            ws[3, start_col_num].color = [255, 185, 15]

        data_txt = pd.read_table(txt_path + '\\' + i)

        #   写入良率信息数据
        ws[index_series[0], start_col_num].options(transpose=True).value = data_txt.iloc[0:9, 1].tolist()

        index_failure_name = data_txt.iloc[:, 0].tolist().index('Failure Detail BreakDown:')
        index_retest_name = data_txt.iloc[:, 0].tolist().index('Retest Detail BreakDown:')
        data_txt_rows = len(data_txt)

        #   写入Fail数据
        ws[index_series[1], start_col_num].value = file_matrix[i][2]  # Overlay 写入
        if ws[index_series[1] + 1, 2].value != None:
            for j in range(index_failure_name + 2, index_retest_name - 1, 1):
                if data_txt.iloc[j, 1] in ws[index_series[1], 2].expand('down').value:
                    ws[ws[index_series[1], 2].expand('down').value.index(data_txt.iloc[j, 1]) + index_series[
                        1], start_col_num].value = data_txt.iloc[j, 3]
                else:
                    insert_row_index = ws[index_series[1], 2].expand('down').last_cell.row
                    ws.api.Rows(insert_row_index + 1).Insert()
                    ws[insert_row_index, 1].value = ws[insert_row_index - 1, 1].value + 1
                    ws[insert_row_index, 2].value = data_txt.iloc[j, 1]
                    ws[insert_row_index, start_col_num].value = data_txt.iloc[j, 3]
        else:
            for n in range(index_retest_name - index_failure_name - 3):
                ws.api.Rows(index_series[1] + 3).Insert()

            ws[index_series[1] + 1, 2].options(transpose=True).value = data_txt.iloc[
                                                                       (index_failure_name + 2):(index_retest_name - 1),
                                                                       1].tolist()
            ws[index_series[1] + 1, 1].options(transpose=True).value = list(
                range(1, index_retest_name - index_failure_name - 2, 1))
            ws[index_series[1] + 1, start_col_num].options(transpose=True).value = data_txt.iloc[
                                                                                   (index_failure_name + 2):(
                                                                                           index_retest_name - 1),
                                                                                   3].tolist()

        #   写入Retest数据，写入Retest数据前要更新Index，因为Failure数据可能影响index

        index_series_data = pd.Series(ws.range('B1:B400').value)
        index_series = index_series_data[index_series_data.values == file_matrix[i][
            1]].index.tolist()  # station位置index，第一个是良率表格，第二个是Result，第三个是Retest

        ws[index_series[2], start_col_num].value = file_matrix[i][2]  # Overlay 写入
        if ws[index_series[2] + 1, 2].value != None:
            for x in range(index_retest_name + 2, data_txt_rows, 1):
                if data_txt.iloc[x, 1] in ws[index_series[2], 2].expand('down').value:
                    ws[ws[index_series[2], 2].expand('down').value.index(data_txt.iloc[x, 1]) + index_series[
                        2], start_col_num].value = data_txt.iloc[x, 3]
                else:
                    insert_row_index = ws[index_series[2], 2].expand('down').last_cell.row
                    ws.api.Rows(insert_row_index + 1).Insert()
                    ws[insert_row_index, 1].value = ws[insert_row_index - 1, 1].value + 1
                    ws[insert_row_index, 2].value = data_txt.iloc[x, 1]
                    ws[insert_row_index, start_col_num].value = data_txt.iloc[x, 3]
        else:
            for n in range(data_txt_rows - index_retest_name - 2):
                ws.api.Rows(index_series[2] + 3).Insert()

            ws[index_series[2] + 1, 2].options(transpose=True).value = data_txt.iloc[(index_retest_name + 2):,
                                                                       1].tolist()
            ws[index_series[2] + 1, 1].options(transpose=True).value = list(
                range(1, data_txt_rows - index_retest_name - 1, 1))
            ws[index_series[2] + 1, start_col_num].options(transpose=True).value = data_txt.iloc[
                                                                                   (index_retest_name + 2):, 3].tolist()

        print_info = f'Finished -> {i}'
        txt_print(text_info, '', print_info, 50, 'Cyan1', 'Times', 15)

        # config标签，判断是否到了一个config的结束位置
        config_flag = 0
        if (file_matrix.columns.tolist().index(i) + 1) == len(file_matrix.columns):
            config_flag = 1
        elif file_matrix[i][0] != file_matrix.iloc[0, file_matrix.columns.tolist().index(i) + 1]:
            config_flag = 1

        # 做两件事情，一个是进行Config文件打包，另外形成Total列，汇总config。
        if config_flag:
            save_name_list = os.path.split(template_path)[-1].split('_')
            zip_name = save_name_list[0] + '_' + save_name_list[1] + '_' + file_matrix[i][0] + '_' + now_date_2 + '.zip'
            z = zipfile.ZipFile(os.path.join(txt_path, zip_name), 'w', zipfile.ZIP_DEFLATED)
            for file_name in config_file_list:
                file_name = str(file_name).replace('txt', 'csv')
                z.write(os.path.join(txt_path, file_name), arcname=file_name)
                config_file_list = []
            z.close()

            old_config_name = pd.Series(ws[1, 0:(start_col_num + 1)].value)
            # if file_matrix[i][0] in old_config_name.values:
            if old_config_name.values.tolist().count(file_matrix[i][0]) > 1:
                old_config_index = old_config_name[old_config_name.values == file_matrix[i][0]].index.tolist()
                start_col_num = start_col_num + 1
                ws.api.Columns(start_col_num + 1).Insert()
                ws[1, start_col_num].value = file_matrix[i][0]
                ws[2, start_col_num].value = 'Total'
                ws[3, start_col_num].value = file_matrix[i][0]
                ws[1:4, start_col_num].color = [0, 140, 255]

                total_flag = 0
                for kk in old_config_index:
                    if ws[2, kk].value == 'Total':
                        old_config_index.remove(kk)
                        total_flag = 1
                        break

                station_unique = file_matrix.iloc[1, :].unique().tolist()
                for ii in station_unique:
                    index_series_total = index_series_data[index_series_data.values == ii].index.tolist()

                    TQ = 0
                    FQ = 0
                    FPQ = 0
                    RQ = 0
                    FLQ = 0

                    for jj in old_config_index:
                        TQ = TQ + ws[index_series_total[0], jj].value
                        FQ = FQ + ws[index_series_total[0] + 1, jj].value
                        FPQ = FPQ + ws[index_series_total[0] + 3, jj].value
                        RQ = RQ + ws[index_series_total[0] + 5, jj].value
                        FLQ = FLQ + ws[index_series_total[0] + 7, jj].value

                    ws[index_series_total[0], start_col_num].value = TQ
                    ws[index_series_total[0] + 1, start_col_num].value = FQ
                    ws[index_series_total[0] + 2, start_col_num].value = FQ / TQ
                    ws[index_series_total[0] + 3, start_col_num].value = FPQ
                    ws[index_series_total[0] + 4, start_col_num].value = FPQ / TQ
                    ws[index_series_total[0] + 5, start_col_num].value = RQ
                    ws[index_series_total[0] + 6, start_col_num].value = RQ / TQ
                    ws[index_series_total[0] + 7, start_col_num].value = FLQ
                    ws[index_series_total[0] + 8, start_col_num].value = FLQ / TQ

                    ws[index_series_total[1], start_col_num].value = ws[
                        index_series_total[1], start_col_num - 1].value  # Fail Overlay
                    for aa in range(ws[index_series_total[1] + 1, 2].expand('down').size):
                        TTF = 0
                        for bb in old_config_index:
                            if ws[index_series_total[1] + 1 + aa, bb].value != None:
                                TTF = TTF + ws[index_series_total[1] + 1 + aa, bb].value * ws[
                                    index_series_total[0], bb].value
                        if TTF != 0:
                            ws[index_series_total[1] + 1 + aa, start_col_num].value = TTF / TQ

                    ws[index_series_total[2], start_col_num].value = ws[
                        index_series_total[1], start_col_num - 1].value  # Fail Overlay
                    for aa in range(ws[index_series_total[2] + 1, 2].expand('down').size):
                        TTRF = 0
                        for bb in old_config_index:
                            if ws[index_series_total[2] + 1 + aa, bb].value != None:
                                TTRF = TTRF + ws[index_series_total[2] + 1 + aa, bb].value * ws[
                                    index_series_total[0], bb].value
                        if TTRF != 0:
                            ws[index_series_total[2] + 1 + aa, start_col_num].value = TTRF / TQ

                if total_flag:
                    ws.api.Columns(kk + 1).Delete()
                    start_col_num = start_col_num - 1

                print_info = f'Finished -> {i} - Total'
                txt_print(text_info, '', print_info, 50, 'Cyan1', 'Times', 15)

        #   file_last = file_matrix[i]

    # save_name_list = os.path.split(template_path)[-1].split('_')
    save_name = save_name_list[0] + '_' + save_name_list[1] + '_' + save_name_list[2] + '_' + now_date_2 + '.' + save_name_list[-1].split('.')[-1]
    wb.save(os.path.join(os.path.split(template_path)[0], save_name))
    app.quit()
    return save_name

# txt_report(template_path, txt_path)
