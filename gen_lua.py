# -*- coding: utf-8 -*-
import argparse
import os
import xlrd
import excel2lua

if __name__ =='__main__':
    parser = argparse.ArgumentParser(description='manual to this script')

    parser.add_argument('--excel_path', type = str, required = True)
    parser.add_argument('--out_path', type = str, required = True)
    parser.add_argument('--classify', type = int, default = 3)
    args = parser.parse_args()
    excel_path = args.excel_path
    out_path = args.out_path
    classify = args.classify

    if os.path.exists(out_path):
        os.system('rmdir /s/q %s'%(out_path))
    os.mkdir(out_path)

    for file in os.listdir(excel_path):
        if file.endswith(".xlsx"):
            print(excel_path + os.sep  + file)
            excel_obj = xlrd.open_workbook(excel_path + os.sep + file)
            filename = os.path.splitext(file)[0]

            if filename == "CONST":
                lua_str = excel2lua.gen_const(excel_obj.sheets()[0])
                if len(lua_str) != 0:
                    lua_table = 'CONST = {\n%s}\nreturn CONST'%(lua_str)
                    with open(os.path.join(out_path, 'const.lua'), "w", encoding='utf-8') as f:
                        f.writelines(lua_table)
                        f.close()
            elif filename == "ERROR":
                lua_str = excel2lua.gen_error(excel_obj.sheets()[0])
                if len(lua_str) != 0:
                    lua_table = 'ERROR = {\n%s}\nreturn ERROR'%(lua_str)
                    with open(os.path.join(out_path, 'error.lua'), "w", encoding='utf-8') as f:
                        f.writelines(lua_table)
                        f.close()
            else:
                lua_str = excel2lua.gen_data(filename, excel_obj.sheets()[0], classify)
                if len(lua_str) != 0:
                    lua_table = 'data_%s = {\n%s}\nreturn data_%s'%(filename, lua_str, filename)
                    with open(os.path.join(out_path, 'data_%s.lua'%(filename)), "w", encoding='utf-8') as f:
                        f.writelines(lua_table)
                        f.close()






