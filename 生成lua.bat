@echo off

echo generate data start
python gen_lua.py --excel_path "C:\Users\dxb\Desktop\excel2lua\excel"  --out_path "C:\Users\dxb\Desktop\excel2lua\data" --classify 3
echo generate data end

pause