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
        第3行: 类型(bool,int,float,string)，不填则代表nil
        第4行: 注释
        第5行: 分类(1:全部 2:客户端 3:服务端)

        第1列: 保留(*代表这一行不读取)
'''
def gen_data(filename, sheet, classify):
    field_row = 1                   # 字段名行
    type_row = 2                    # 类型行
    classify_row = 4                # 分类行
    start_row = 5                   # 读表起始行
    start_col = 2                   # 字段起始列
    l_fisrl_row = sheet.row_values(0)                   # 第一行值的列表
    l_type_row = sheet.row_values(type_row)             # 字段类型值的列表
    l_field_row = sheet.row_values(field_row)           # 字段名值的列表
    l_classify_row = sheet.row_values(classify_row)     # 分类值的列表
    l_first_col = sheet.col_values(0)                   # 第一列值的列表

    lua_str = ''
    for i in range(start_row, sheet.nrows): #逐行
        if l_first_col[i] != "*":
            row_value = sheet.row_values(i) #该行数据
            # 一行数据开头
            line = '\t[%d] =\n\t\t{\n'%(int(row_value[1]))
            for j in range(start_col, sheet.ncols):
                if l_fisrl_row[j] == "*":
                    pass
                elif l_classify_row[j] != 1 and l_classify_row[j] != classify:
                    pass
                else:
                    cell_type = sheet.cell(i, j).ctype
                    cell_value = sheet.cell(i, j).value
                    if cell_type != 0:              # 此行该字段为空的舍弃
                        tips = ""
                        if l_type_row[j] == "string":
                            if cell_type == 1:
                                if cell_value.isspace() == True:
                                    tips = '%s空格字符串'%(l_field_row[j])
                                else:
                                    line += '\t\t%s = "%s",\n'%(l_field_row[j], cell_value)
                            else:
                                tips = '%s格式应为string'%(l_field_row[j])

                        elif l_type_row[j] == "int":
                            if cell_type == 2 and cell_value % 1 == 0:
                                line += '\t\t%s = %d,\n'%(l_field_row[j], int(cell_value))
                            else:
                                tips = '%s格式应为int'%(l_field_row[j])

                        elif l_type_row[j] == "float":
                            if cell_type == 2 and cell_value % 1 != 0:
                                line += '\t\t%s = %f,\n'%(l_field_row[j], float(cell_value))
                            else:
                                tips = '%s格式应为float'%(l_field_row[j])
                        elif l_type_row[j] == "bool":
                            if cell_value == "true" or cell_value == "false":
                                line += '\t\t%s = %s,\n'%(l_field_row[j], cell_value)
                            else:
                                tips = '%s格式应为bool'%(l_field_row[j])

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
                if type_value== "string":
                    if value_type == 1:
                        if value.isspace() == True:
                            tips = "是空格字符串"
                        else:
                            value = '"%s"' % value
                    else:
                        tips = "格式应为string"
                elif type_value == "int":
                    if value_type == 2 and value % 1 == 0:
                        value = int(value)
                    else:
                        tips = "格式应为int"
                elif type_value == "float":
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
                raise Exception("const表%s%s" % (sheet.cell(i, field_col).value, tips))

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