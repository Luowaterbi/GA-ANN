import xlrd
from random import shuffle


def get_content(inpath):
    data = xlrd.open_workbook(inpath, encoding_override='utf-8')
    table = data.sheets()[0]
    ncols = table.ncols
    cols_info = []
    for i in range(12, 76):
        all_data = table.row_values(i)
        result = []
        for j in range(3, ncols - 2):
            if all_data[j] == '':
                continue
            if isinstance(all_data[j], str):
                result.append(float(all_data[j][:-3]))
            else:
                result.append(all_data[j])
        if result != []:
            cols_info.append(result)
    return cols_info


if __name__ == "__main__":
    info = get_content("Appendix Data for Revisiting H-R Rules.xls")
    fin = open('input.txt', 'w')
    fout = open('output.txt', 'w')
    shuffle(info)
    for i in info:
        fout.write("%f\n" % (i[0]))
        print(i[0])
        fin.write("%f %f %f %f\n" %
                  ((i[1] - i[2]) / i[1], i[3], i[4], i[5] - i[6]))
    fin.close()
    fout.close()
