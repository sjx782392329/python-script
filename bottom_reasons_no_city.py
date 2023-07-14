# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import xlrd
import xlsxwriter


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


def read_excel(sheet, type):
    rows = sheet.nrows
    cols = sheet.ncols
    #八个问题
    lists = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    count = 0
    for i in range(1, rows):
        lines = sheet.row_values(i, 0, cols)
        if type == lines[0]:
            count = count + 1
            # 问题是从第5列开始
            tmp_list = lines[4:cols]
            # 8个问题
            for j in range(0, 8):
                 lists[j] = lists[j] + tmp_list[j]
    # 8个问题
    percent_lists = [] * 8
    for l in lists:
        if count == 0.0 or count == 0:
            percent_lists.append(-1)
        else:
            a = l / count
            percent_lists.append('{:.1%}'.format(a))

    res_list = [type, lists, count, percent_lists]
    return res_list


def write_excel(workbook, final_lists):
    problem_list = ['不满意原因', '房源图片不全面', '家具家电信息不详细', '缺少VR', '缺少户型图', '小区信息不完善', '周边配套信息不完善', '租金收费信息不清晰', '其他']
    write_row = 0
    write_col = 1
    write_work_sheet = workbook.get_worksheet_by_name('test')
    for problem in problem_list:
                write_work_sheet.write(write_row + 1, 0, problem)
                write_row = write_row + 1

    for lists in final_lists:
        write_row = 0
        # type
        write_work_sheet.write(write_row, write_col, lists[0])
        write_row = write_row + 1
        write_work_sheet.write(write_row, write_col, '计数')
        write_work_sheet.write(write_row, write_col + 1, '占比')
        write_row = write_row + 1
        # 8个问题
        for i in range(0, 8):
            # lists
            write_work_sheet.write(write_row, write_col, lists[1][i])
            # percent
            write_work_sheet.write(write_row, write_col + 1, lists[3][i])
            write_row = write_row + 1
            # count
        write_work_sheet.write(write_row, write_col, lists[2])
        write_col = write_col + 2
        write_row = 0


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    table = xlrd.open_workbook("/Users/shenjinxin/Desktop/bottom_reasons.xlsx")
    read_sheet = table.sheet_by_name("sheet1")

    write_workbook = xlsxwriter.Workbook("/Users/shenjinxin/Desktop/afterbottom_reasons.xlsx")
    # 创建一个名字为 test 的 excel
    write_workbook.add_worksheet("test")

    type_list = ['house', 'qingtuoguan', 'fensanshigongyu', 'jizhongshigongyu']
    final_list = []
    for type in type_list:
        res = read_excel(read_sheet, type)
        final_list.append(res)
    print(final_list)
    write_excel(write_workbook, final_list)
    write_workbook.close()



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
