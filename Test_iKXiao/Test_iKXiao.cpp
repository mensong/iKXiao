// Test_iKXiao.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include "..\iKXiao\iKXiao.h"

void normalTest()
{
    WorkBook* wb = iKXiao::Ins().OpenExcel("admin.xlsx", NULL);
    if (wb)
    {
        WorkSheet* sheet = wb->OpenCurrentSheet();
        if (sheet)
        {
            size_t rowCount = sheet->GetRowCount();
            for (size_t r = 0; r < rowCount; r++)
            {
                size_t columnCount = 0;
                char* rowData = sheet->GetRowStringArray(r, &columnCount);
                size_t offset = 0;
                for (size_t c = 0; c < columnCount; c++)
                {
                    std::cout << rowData + offset << ",";
                    offset += strlen(rowData + offset) + 1;
                }
                sheet->FreeString(rowData);
                std::cout << std::endl;
            }

            Cell* cell = sheet->OpenCell(0, 0);
            if (cell)
            {
                cell->SetStringValue("我爱你mensong");
                std::string cellstr = cell->GetStringValue("");
                std::cout << "(0,0) = " << cellstr << std::endl;
                sheet->CloseCell(cell);
            }


            cell = sheet->OpenCell("A1");
            if (cell)
            {
                std::string cellstr = cell->GetStringValue("");
                std::cout << "A1 = " << cellstr << std::endl;
                sheet->CloseCell(cell);
            }

            cell = sheet->OpenCell("A2");
            if (cell)
            {
                cell->SetFormula("=SUM(1,2,3)");
                int intVal = cell->GetIntValue(0);
                std::cout << "=SUM(1,2,3) = " << intVal << std::endl;
                sheet->CloseCell(cell);
            }

            wb->CloseSheet(sheet);
        }

        wb->Save("admin1.xlsx");

        iKXiao::Ins().CloseExcel(wb);
    }
}

void testConfig()
{
    WorkBook* wb = iKXiao::Ins().OpenExcel("config.xlsx", NULL);
    if (wb)
    {
        WorkSheet* sheet = wb->OpenCurrentSheet();
        if (sheet)
        {            
            size_t idxRowStart = sheet->GetNotEmptyRowStart();
            size_t idxRowEnd = sheet->GetNotEmptyRowEnd();
            size_t idxColumnStart = sheet->GetNotEmptyColumnStart();
            size_t idxColumnEnd = sheet->GetNotEmptyColumnEnd();
            for (size_t r = idxRowStart; r <= idxRowEnd; r++)
            {
                for (size_t c = idxColumnStart; c <= idxColumnEnd; c++)
                {
                    Cell* cell = sheet->OpenCell(r, c);
                    if (cell)
                    {
                        auto type = cell->GetValueType();
                        switch (type)
                        {
                        case Cell::empty_value:
                            std::cout << "<NULL>" << ',';
                            break;
                        case Cell::boolean_value:
                            std::cout << cell->GetBoolValue() << ',';
                            break;
                        case Cell::date_value:
                        {
                            int year = 0;
                            int month = 0;
                            int day = 0;
                            cell->GetDateValue(year, month, day);
                            std::cout << year << "年" << month << "月" << day << "日" << ',';
                            break;
                        }
                        case Cell::error_value:
                            std::cout << "#VALUE!" << ',';
                            break;
                        case Cell::inline_string_value:
                        case Cell::shared_string_value:
                            std::cout << cell->GetStringValue() << ',';
                            break;
                        case Cell::number_value:
                            std::cout << cell->GetDoubleValue() << ',';
                            break;
                        case Cell::formula_string_value:
                            std::cout << cell->GetDoubleValue() << ',';
                            break;
                        default:
                            std::cout << cell->GetStringValue() << ',';
                            break;
                        }

                        sheet->CloseCell(cell);
                    }
                }
                std::cout << std::endl;
            }

            wb->CloseSheet(sheet);
        }

        iKXiao::Ins().CloseExcel(wb);
    }
}

int main()
{
    normalTest();

    std::cout << std::endl << std::endl;
    
    testConfig();

    return 0;
}
