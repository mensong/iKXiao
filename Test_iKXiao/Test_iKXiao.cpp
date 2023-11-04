// Test_iKXiao.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include "..\iKXiao\iKXiao.h"

int main()
{
    IK_WORKBOOK wb = iKXiao::Ins().OpenExcel("E:\\admin.xlsx", NULL);

    size_t rowCount = iKXiao::Ins().GetRowCount(wb, IDX_SHEET_CUR);
    for (size_t r = 0; r < rowCount; r++)
    {
        size_t columnCount = 0;
        char* rowData = iKXiao::Ins().GetRowStringArray(wb, IDX_SHEET_CUR, r, &columnCount);
        size_t offset = 0;
        for (size_t c = 0; c < columnCount; c++)
        {
            std::cout << rowData + offset << ",";
            offset += strlen(rowData + offset) + 1;
        }
        iKXiao::Ins().FreeString(wb, rowData);
        std::cout << std::endl;
    }
    

    char* cellStr = iKXiao::Ins().GetCellStringValue(wb, IDX_SHEET_CUR, 0, 0);
    char* cellStr1 = iKXiao::Ins().GetCellStringValueByRefName(wb, IDX_SHEET_CUR, "A1");

    iKXiao::Ins().SetCellStringValue(wb, IDX_SHEET_CUR, 0, 0, "我爱你mensong");
    iKXiao::Ins().SetCellFormula(wb, IDX_SHEET_CUR, 0, 1, "=SUM(1,2,3)");
    auto type = iKXiao::Ins().GetCellType(wb, IDX_SHEET_CUR, 0, 1);
    char* formulaValue = iKXiao::Ins().GetCellStringValue(wb, IDX_SHEET_CUR, 0, 1);

    iKXiao::Ins().SaveExcel(wb, "E:\\admin1.xlsx", NULL);

    iKXiao::Ins().CloseExcel(wb);
    return 0;
}
