# -*- coding: utf-8 -*-
''' 
---------------------------------------------
--------------excel生成lua的模块-------------
---------------------------------------------
'''

'''
    brief: excel文件生成数据表
    param:
        filename: 文件名
        sheet: sheet对象
        classify: 分类(1:全部 2:客户端 3:服务端)

    return: lua数据表的字符串

    注意sheet格式:
        第1行: 保留(*代表这一列不读取)
        第2行: 字段名
        第3行: 类型(INT,STRING)
        第4行: 注释
        第5行: classify类型值

        第1列: 保留(*代表这一行不读取)
'''
def gen_data(filename, sheet, classify):
    field_row = 1                   # 字段名行
    type_row = 2                    # 类型行
    start_row = 6                   # 读表起始行
    start_col = 2                   # 读表起始列
    l_fisrl_row = sheet.row_values(0)                   # 第一行值的列表
    l_first_col = sheet.col_values(0)                   # 第一列值的列表
    l_type_row = sheet.row_values(type_row)             # 字段类型值的列表
    l_field_row = sheet.row_values(field_row)           # 字段名值的列表

    lua_str = ''
    for i in range(start_row, sheet.nrows): #逐行
        if l_first_col[i] != "*":
            row_value = sheet.row_values(i) #该行数据

            # 一行数据开头
            line = '\t[%d] =\n\t\t{\n'%(int(row_value[1]))
            for j in range(start_col, sheet.ncols):
                if l_fisrl_row[j] != "*":
                    cell_type = sheet.cell(i, j).ctype
                    cell_value = sheet.cell(i, j).value
                    if cell_type != 0:              # 此行该字段为空的舍弃
                        tips = ""
                        if l_type_row[j] == "STRING":
                            if cell_type == 1:
                                if cell_value.isspace() == True:
                                    tips = "空格字符串"
                                else:
                                    line += '\t\t%s = "%s",\n'%(l_field_row[j], cell_value)
                            else:
                                tips = "格式应为string"
                        elif l_type_row[j] == "INT":
                            if cell_type == 2 and cell_value % 1 == 0:
                                line += '\t\t%s = %d,\n'%(l_field_row[j], int(cell_value))
                            else:
                                tips = "格式应为int"

                        elif l_type_row[j] == "FLOAT":
                            if cell_type == 2 and cell_value % 1 != 0:
                                line += '\t\t%s = %f,\n'%(l_field_row[j], float(cell_value))
                            else:
                                tips = "格式应为float"

                        if len(tips) != 0:
                            raise Exception('%s表的第%d行,%d列, %s'%(filename, i+1, j+1, tips))

            # 一行数据结尾
            line += "\t},\n"
            lua_str += line

    return lua_str


'''
    brief: excel文件生成常量表
    param:
        sheet: sheet对象

    return: lua数据表的字符串

    注意sheet格式:
        第1行: 保留(*代表这一列不读取)

        第1列: 保留(*代表这一行不读取)
        第2列: 字段名
        第3列: 字段值
        第4行: 注释
        第5行: 类型(INT,STRING)
'''
def gen_const(sheet):
    field_col = 1                                       # 字段名列
    value_col = 2                                       # 字段值列
    annotate_col = 3                                    # 注释列
    type_col = 4                                        # 字段类型列
    align_len = 52                                      # 注释对齐行

    l_first_col = sheet.col_values(0)                   # 第一列值的列表
    lua_str = ''
    for i in range(1, sheet.nrows): #逐行
        if l_first_col[i] != "*":
            tips = ""
            value = sheet.cell(i, value_col).value
            value_type = sheet.cell(i, value_col).ctype
            if value_type != 0:              # 此行该字段为空的舍弃
                type_value = sheet.cell(i, type_col).value 
                if type_value== "STRING":
                    if value_type == 1:
                        if value.isspace() == True:
                            tips = "是空格字符串"
                        else:
                            value = '"%s"' % value
                    else:
                        tips = "格式应为string"
                elif type_value == "INT":
                    if value_type == 2 and value % 1 == 0:
                        value = int(value)
                    else:
                        tips = "格式应为int"
                elif type_value == "FLOAT":
                    if value_type == 2 and value % 1 != 0:
                        value = float(value)
                    else:
                        tips = "格式应为float"

                temp = '    %s = %s,'%(sheet.cell(i, field_col).value, value)
                space = ' ' * (align_len - len(temp))
                temp = '%s%s--%s\n' % (temp, space, sheet.cell(i, annotate_col).value)
                lua_str += temp
            else:
                tips = "没有填值"

            if len(tips) != 0:
                raise Exception("%s%s" % (sheet.cell(i, field_col).value, tips))

    return lua_str

'''
    brief: excel文件生成错误码表
    param:
        sheet: sheet对象

    return: lua数据表的字符串

    注意sheet格式:
        第1行: 保留(*代表这一列不读取)

        第1列: 保留(*代表这一行不读取)
        第2列: 字段名
        第3列: 字段值
        第4行: 注释
'''
def gen_error(sheet):
    field_col = 1                                       # 字段名列
    value_col = 2                                       # 字段值列
    annotate_col = 3                                    # 注释列
    align_len = 52                                      # 注释对齐行

    l_first_col = sheet.col_values(0)                   # 第一列值的列表
    lua_str = ''
    for i in range(1, sheet.nrows): #逐行
        if l_first_col[i] != "*":
            tips = ""
            value = sheet.cell(i, value_col).value
            value_type = sheet.cell(i, value_col).ctype
            if value_type != 0:              # 此行该字段为空的舍弃
                if value_type == 2 and value % 1 == 0:
                    value = int(value)
                else:
                    tips = "格式应为int"

                temp = '    %s = %s,'%(sheet.cell(i, field_col).value, value)
                space = ' ' * (align_len - len(temp))
                temp = '%s%s--%s\n' % (temp, space, sheet.cell(i, annotate_col).value)
                lua_str += temp
            else:
                tips = "没有填值"

            if len(tips) != 0:
                raise Exception("%s%s" % (sheet.cell(i, field_col).value, tips))

    return lua_str