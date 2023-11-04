#pragma once
#ifndef _AFX
#include <windows.h>
#endif
#include <string>
#include <stdint.h>

#ifdef IKXIAO_EXPORTS
#define IKXIAO_API extern "C" __declspec(dllexport)
#else
#define IKXIAO_API extern "C" __declspec(dllimport)
#endif

typedef void* IK_WORKBOOK;
//typedef void* IK_WORKSHEET;

#define IDX_SHEET_CUR -1
#define IDX_SHEET_LAST -2

enum CellType
{
	/// no value
	empty_value,
	/// value is TRUE or FALSE
	boolean_value,
	/// value is an ISO 8601 formatted date
	date_value,
	/// value is a known error code such as \#VALUE!
	error_value,
	/// value is a string stored in the cell
	inline_string_value,
	/// value is a number
	number_value,
	/// value is a string shared with other cells to save space
	shared_string_value,
	/// value is the string result of a formula
	formula_string_value
};

IKXIAO_API IK_WORKBOOK	OpenExcel(const char* xlsxFilepath, const char* password);
IKXIAO_API int			GetSheetCount(IK_WORKBOOK workbook);
IKXIAO_API int			GetSheetIndexByTitle(IK_WORKBOOK workbook, const char* sheetTitle);
IKXIAO_API char*		GetSheetTitle(IK_WORKBOOK workbook, int sheetIndex);
IKXIAO_API bool			SetCurrentSheet(IK_WORKBOOK workbook, int sheetIndex);
IKXIAO_API int			CreateSheet(IK_WORKBOOK workbook, int atIndex = -1);
IKXIAO_API int			CopySheet(IK_WORKBOOK workbook, int srcIndex, int atIndex = -1);
IKXIAO_API bool			RemoveSheet(IK_WORKBOOK workbook, int sheetIndex);
IKXIAO_API bool			SetSheetTitle(IK_WORKBOOK workbook, int sheetIndex, const char* title);
IKXIAO_API size_t		GetRowCount(IK_WORKBOOK workbook, int sheetIndex);
//获得行数据，数据格式为"列1内容\0列2内容\0列3内容\0"
IKXIAO_API char*		GetRowStringArray(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t* columnCount);
IKXIAO_API CellType		GetCellType(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex);
IKXIAO_API char*		GetCellStringValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex);
IKXIAO_API char*		GetCellStringValueByRefName(IK_WORKBOOK workbook, int sheetIndex, const char* refName);
IKXIAO_API bool			SetCellStringValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, const char* strVal);
IKXIAO_API bool			SetCellNullValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex);
IKXIAO_API bool			SetCellBoolValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, bool boolean_value);
IKXIAO_API bool			GetCellBoolValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, bool defVal = false);
IKXIAO_API bool			SetCellIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int int_value);
IKXIAO_API int			GetCellIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int defVal = 0);
IKXIAO_API bool			SetCellUIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int int_value);
IKXIAO_API unsigned int GetCellUIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int defVal = 0);
IKXIAO_API bool			SetCellLLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, long long int int_value);
IKXIAO_API long long int GetCellLLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, long long int defVal = 0);
IKXIAO_API bool			SetCellULLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int int_value);
IKXIAO_API unsigned long long int GetCellULLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int defVal = 0);
IKXIAO_API bool			SetCellDoubleValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, double float_value);
IKXIAO_API double		GetCellDoubleValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, double defVal = 0.0);
IKXIAO_API bool			SetCellDateValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int year_, int month_, int day_);
IKXIAO_API bool			GetCellDateValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& year_, int& month_, int& day_);
IKXIAO_API bool			SetCellDatetimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0);
IKXIAO_API bool			GetCellDatetimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& year_, int& month_, int& day_, int& hour_, int& minute_, int& second_, int& microsecond_);
IKXIAO_API bool			SetCellTimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0);
IKXIAO_API bool			GetCellTimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& hour_, int& minute_, int& second_, int& microsecond_);
IKXIAO_API bool			SetCellTimeDeltaValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int days_, int hours_, int minutes_, int seconds_, int microseconds_);
IKXIAO_API bool			GetCellTimeDeltaValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int days_, int& hours_, int& minutes_, int& seconds_, int& microseconds_);
IKXIAO_API bool			SetCellFormula(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, const char* formula);
IKXIAO_API bool			CopyCell(IK_WORKBOOK workbook, int sheetIndex, size_t fromRowIndex, size_t fromColumnIndex, size_t toRowIndex, size_t toColumnIndex);
IKXIAO_API bool			SaveExcel(IK_WORKBOOK workbook, const char* filepath, const char* password);
IKXIAO_API void			FreeString(IK_WORKBOOK workbook, char* buff);
IKXIAO_API void			CloseExcel(IK_WORKBOOK workbook);

class iKXiao
{
#define DEF_PROC(name) \
	decltype(::name)* name

#define SET_PROC(hDll, name) \
	this->name = (decltype(::name)*)::GetProcAddress(hDll, #name)

public:
	iKXiao()
	{
		hDll = LoadLibraryFromCurrentDir("iKXiao.dll");
		if (!hDll)
			return;

		SET_PROC(hDll, OpenExcel);
		SET_PROC(hDll, GetSheetCount);
		SET_PROC(hDll, GetSheetIndexByTitle);
		SET_PROC(hDll, GetSheetTitle);
		SET_PROC(hDll, SetCurrentSheet);
		SET_PROC(hDll, CreateSheet);
		SET_PROC(hDll, CopySheet);
		SET_PROC(hDll, RemoveSheet);
		SET_PROC(hDll, SetSheetTitle);
		SET_PROC(hDll, GetRowCount);
		SET_PROC(hDll, GetRowStringArray);
		SET_PROC(hDll, GetCellType);
		SET_PROC(hDll, GetCellStringValue);
		SET_PROC(hDll, GetCellStringValueByRefName);
		SET_PROC(hDll, SetCellStringValue);
		SET_PROC(hDll, SetCellNullValue);
		SET_PROC(hDll, SetCellBoolValue);
		SET_PROC(hDll, GetCellBoolValue);
		SET_PROC(hDll, SetCellIntValue);
		SET_PROC(hDll, GetCellIntValue);
		SET_PROC(hDll, SetCellUIntValue);
		SET_PROC(hDll, GetCellUIntValue);
		SET_PROC(hDll, SetCellLLIntValue);
		SET_PROC(hDll, GetCellLLIntValue);
		SET_PROC(hDll, SetCellULLIntValue);
		SET_PROC(hDll, GetCellULLIntValue);
		SET_PROC(hDll, SetCellDoubleValue);
		SET_PROC(hDll, GetCellDoubleValue);
		SET_PROC(hDll, SetCellDateValue);
		SET_PROC(hDll, GetCellDateValue);
		SET_PROC(hDll, SetCellDatetimeValue);
		SET_PROC(hDll, GetCellDatetimeValue);
		SET_PROC(hDll, SetCellTimeValue);
		SET_PROC(hDll, GetCellTimeValue);
		SET_PROC(hDll, SetCellTimeDeltaValue);
		SET_PROC(hDll, GetCellTimeDeltaValue);
		SET_PROC(hDll, SetCellFormula);
		SET_PROC(hDll, CopyCell);
		SET_PROC(hDll, SaveExcel);
		SET_PROC(hDll, FreeString);
		SET_PROC(hDll, CloseExcel);
	}


	DEF_PROC(OpenExcel);
	DEF_PROC(GetSheetCount);
	DEF_PROC(GetSheetIndexByTitle);
	DEF_PROC(GetSheetTitle);
	DEF_PROC(SetCurrentSheet);
	DEF_PROC(CreateSheet);
	DEF_PROC(CopySheet);
	DEF_PROC(RemoveSheet);
	DEF_PROC(SetSheetTitle);
	DEF_PROC(GetRowCount);
	DEF_PROC(GetRowStringArray);
	DEF_PROC(GetCellType);
	DEF_PROC(GetCellStringValue); 
	DEF_PROC(GetCellStringValueByRefName);
	DEF_PROC(SetCellStringValue);
	DEF_PROC(SetCellNullValue);
	DEF_PROC(SetCellBoolValue);
	DEF_PROC(GetCellBoolValue);
	DEF_PROC(SetCellIntValue);
	DEF_PROC(GetCellIntValue);
	DEF_PROC(SetCellUIntValue);
	DEF_PROC(GetCellUIntValue);
	DEF_PROC(SetCellLLIntValue);
	DEF_PROC(GetCellLLIntValue);
	DEF_PROC(SetCellULLIntValue);
	DEF_PROC(GetCellULLIntValue);
	DEF_PROC(SetCellDoubleValue);
	DEF_PROC(GetCellDoubleValue);
	DEF_PROC(SetCellDateValue);
	DEF_PROC(GetCellDateValue);
	DEF_PROC(SetCellDatetimeValue);
	DEF_PROC(GetCellDatetimeValue);
	DEF_PROC(SetCellTimeValue);
	DEF_PROC(GetCellTimeValue);
	DEF_PROC(SetCellTimeDeltaValue);
	DEF_PROC(GetCellTimeDeltaValue);
	DEF_PROC(SetCellFormula);
	DEF_PROC(CopyCell);
	DEF_PROC(SaveExcel);
	DEF_PROC(FreeString);
	DEF_PROC(CloseExcel);


public:
	static iKXiao& Ins()
	{
		static iKXiao s_ins;
		return s_ins;
	}

	static HMODULE LoadLibraryFromCurrentDir(const char* dllName)
	{
		char selfPath[MAX_PATH];
		MEMORY_BASIC_INFORMATION mbi;
		HMODULE hModule = ((::VirtualQuery(LoadLibraryFromCurrentDir, &mbi, sizeof(mbi)) != 0) ? (HMODULE)mbi.AllocationBase : NULL);
		::GetModuleFileNameA(hModule, selfPath, MAX_PATH);
		std::string moduleDir(selfPath);
		size_t idx = moduleDir.find_last_of('\\');
		moduleDir = moduleDir.substr(0, idx);
		std::string modulePath = moduleDir + "\\" + dllName;
		char curDir[MAX_PATH];
		::GetCurrentDirectoryA(MAX_PATH, curDir);
		::SetCurrentDirectoryA(moduleDir.c_str());
		HMODULE hDll = LoadLibraryA(modulePath.c_str());
		::SetCurrentDirectoryA(curDir);
		if (!hDll)
		{
			DWORD err = ::GetLastError();
			char buf[10];
			sprintf_s(buf, "%u", err);
			::MessageBoxA(NULL, ("找不到" + modulePath + "模块:" + buf).c_str(), "找不到模块", MB_OK | MB_ICONERROR);
		}
		return hDll;
	}
	~iKXiao()
	{
		if (hDll)
		{
			FreeLibrary(hDll);
			hDll = NULL;
		}
	}

private:
	HMODULE hDll;
};

