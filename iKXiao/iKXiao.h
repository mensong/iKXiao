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



class Cell
{
public:
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

public:
	virtual CellType		GetValueType() = 0;
	virtual char*			GetStringValue(const char* defVal = "") = 0;
	virtual bool			SetStringValue(const char* strVal) = 0;
	virtual bool			SetNullValue() = 0;
	virtual bool			SetBoolValue(bool boolean_value) = 0;
	virtual bool			GetBoolValue(bool defVal = false) = 0;
	virtual bool			SetIntValue(int int_value) = 0;
	virtual int				GetIntValue(int defVal = 0) = 0;
	virtual bool			SetUIntValue(unsigned int int_value) = 0;
	virtual unsigned int	GetUIntValue(unsigned int defVal = 0) = 0;
	virtual bool			SetLLIntValue(long long int int_value) = 0;
	virtual long long int	GetLLIntValue(long long int defVal = 0) = 0;
	virtual bool			SetULLIntValue(unsigned long long int int_value) = 0;
	virtual unsigned long long int GetULLIntValue(unsigned long long int defVal = 0) = 0;
	virtual bool			SetDoubleValue(double float_value) = 0;
	virtual double			GetDoubleValue(double defVal = 0.0) = 0;
	virtual bool			SeDateValue(int year_, int month_, int day_) = 0;
	virtual bool			GetDateValue(int& year_, int& month_, int& day_) = 0;
	virtual bool			SetDatetimeValue(int year_, int month_, int day_, 
												 int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0) = 0;
	virtual bool			GetDatetimeValue(int& year_, int& month_, int& day_, 
												 int& hour_, int& minute_, int& second_, int& microsecond_) = 0;
	virtual bool			SetTimeValue(int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0) = 0;
	virtual bool			GetTimeValue(int& hour_, int& minute_, int& second_, int& microsecond_) = 0;
	virtual bool			SetTimeDeltaValue(int days_, int hours_, int minutes_, int seconds_, int microseconds_) = 0;
	virtual bool			GetTimeDeltaValue(int days_, int& hours_, int& minutes_, int& seconds_, int& microseconds_) = 0;
	virtual bool			SetFormula(const char* formula) = 0;
	virtual bool			CopyFrom(Cell* otherCell) = 0;

	virtual void			FreeString(char* buff) = 0;
};

class WorkSheet
{
public:
	virtual char*			GetSheetTitle() = 0;
	virtual bool			SetSheetTitle(const char* title) = 0;
	virtual size_t			GetRowCount() = 0;
	virtual size_t			GetColumnCount() = 0;
	virtual size_t			GetNotEmptyRowStart() = 0;
	virtual size_t			GetNotEmptyRowEnd() = 0;
	virtual size_t			GetNotEmptyColumnStart() = 0;
	virtual size_t			GetNotEmptyColumnEnd() = 0;
	//获得行数据，数据格式为"列1内容\0列2内容\0列3内容\0"
	virtual char*			GetRowStringArray(size_t rowIndex, size_t* columnCount) = 0;

	virtual Cell*			OpenCell(size_t rowIdx, int columnIdx) = 0;
	virtual Cell*			OpenCell(const char* refName) = 0;
	virtual void			CloseCell(Cell* cell) = 0;

	virtual void			FreezePanes(Cell* top_left_cell) = 0;
	virtual void			UnfreezePanes() = 0;
	virtual bool			HasFreezePanes() = 0;

	virtual void			FreeString(char* buff) = 0;
};

class WorkBook
{
public:
	virtual int				GetSheetCount() = 0;
	virtual WorkSheet*		OpenSheetByIndex(int sheetIndex) = 0;
	virtual WorkSheet*		OpenSheetByTitle(const char* sheetTitle) = 0;
	virtual bool			SetCurrentSheet(WorkSheet* sheet) = 0;
	virtual WorkSheet*		OpenCurrentSheet() = 0;
	virtual WorkSheet*		CreateSheet(int atIndex = -1) = 0;
	virtual WorkSheet*		CloneSheet(WorkSheet* srcSheet, int atIndex = -1) = 0;
	virtual bool			RemoveSheet(WorkSheet* sheet) = 0;
	virtual bool			Save(const char* filepath) = 0;

	virtual void			CloseSheet(WorkSheet* sheet) = 0;

	virtual void			FreeString(char* buff) = 0;	
};

IKXIAO_API WorkBook*		OpenExcel(const char* xlsxFilepath, const char* password);
IKXIAO_API void				CloseExcel(WorkBook* workbook);

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
		SET_PROC(hDll, CloseExcel);
	}


	DEF_PROC(OpenExcel);
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
		HMODULE hModule = ((::VirtualQuery(LoadLibraryFromCurrentDir, &mbi, sizeof(mbi)) != 0) ?
			(HMODULE)mbi.AllocationBase : NULL);
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

