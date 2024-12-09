#include "pch.h"
#include "iKXiao.h"
#include <xlnt/xlnt.hpp>
#include <io.h>
#include <set>
#include <vector>

static bool IsPathExist(const char* path)
{
	int nRet = _access(path, 0);
	return 0 == nRet || EACCES == nRet;
}


//将Ansi字符转换为Unicode字符串
static std::wstring AnsiToUnicode(const std::string& multiByteStr)
{
	wchar_t* pWideCharStr; //定义返回的宽字符指针
	int nLenOfWideCharStr; //保存宽字符个数，注意不是字节数
	const char* pMultiByteStr = multiByteStr.c_str();
	//获取宽字符的个数
	nLenOfWideCharStr = MultiByteToWideChar(CP_ACP, 0, pMultiByteStr, -1, NULL, 0);
	//获得宽字符指针
	pWideCharStr = (wchar_t*)(HeapAlloc(GetProcessHeap(), 0, nLenOfWideCharStr * sizeof(wchar_t)));
	MultiByteToWideChar(CP_ACP, 0, pMultiByteStr, -1, pWideCharStr, nLenOfWideCharStr);
	//返回
	std::wstring wideByteRet(pWideCharStr, nLenOfWideCharStr);
	//销毁内存中的字符串
	HeapFree(GetProcessHeap(), 0, pWideCharStr);
	return wideByteRet.c_str();
}

//将Unicode字符转换为Ansi字符串
static std::string UnicodeToAnsi(const std::wstring& wideByteStr)
{
	char* pMultiCharStr; //定义返回的多字符指针
	int nLenOfMultiCharStr; //保存多字符个数，注意不是字节数
	const wchar_t* pWideByteStr = wideByteStr.c_str();
	//获取多字符的个数
	nLenOfMultiCharStr = WideCharToMultiByte(CP_ACP, 0, pWideByteStr, -1, NULL, 0, NULL, NULL);
	//获得多字符指针
	pMultiCharStr = (char*)(HeapAlloc(GetProcessHeap(), 0, nLenOfMultiCharStr * sizeof(char)));
	WideCharToMultiByte(CP_ACP, 0, pWideByteStr, -1, pMultiCharStr, nLenOfMultiCharStr, NULL, NULL);
	//返回
	std::string sRet(pMultiCharStr, nLenOfMultiCharStr);
	//销毁内存中的字符串
	HeapFree(GetProcessHeap(), 0, pMultiCharStr);
	return sRet.c_str();
}

static std::string UnicodeToUtf8(const std::wstring& wideByteStr)
{
	int len = WideCharToMultiByte(CP_UTF8, 0, wideByteStr.c_str(), -1, NULL, 0, NULL, NULL);
	char* szUtf8 = new char[len + 1];
	memset(szUtf8, 0, len + 1);
	WideCharToMultiByte(CP_UTF8, 0, wideByteStr.c_str(), -1, szUtf8, len, NULL, NULL);
	std::string s = szUtf8;
	delete[] szUtf8;
	return s.c_str();
}

static std::wstring Utf8ToUnicode(const std::string& utf8Str)
{
	//预转换，得到所需空间的大小;
	int wcsLen = ::MultiByteToWideChar(CP_UTF8, NULL, utf8Str.c_str(), strlen(utf8Str.c_str()), NULL, 0);
	//分配空间要给'\0'留个空间，MultiByteToWideChar不会给'\0'空间
	wchar_t* wszString = new wchar_t[wcsLen + 1];
	//转换
	::MultiByteToWideChar(CP_UTF8, NULL, utf8Str.c_str(), strlen(utf8Str.c_str()), wszString, wcsLen);
	//最后加上'\0'
	wszString[wcsLen] = '\0';
	std::wstring s(wszString);
	delete[] wszString;
	return s;
}

static std::string AnsiToUtf8(const std::string& multiByteStr)
{
	std::wstring ws = AnsiToUnicode(multiByteStr);
	return UnicodeToUtf8(ws);
}

static std::string Utf8ToAnsi(const std::string& utf8Str)
{
	std::wstring ws = Utf8ToUnicode(utf8Str);
	return UnicodeToAnsi(ws);
}

class BufferManager
{
public:
	BufferManager()
	{

	}
	virtual ~BufferManager()
	{
		for (auto it = m_stringCache.begin(); it != m_stringCache.end(); ++it)
		{
			delete[](*it);
		}
		m_stringCache.clear();
	}
public:
	char* allocStringBuffer(int size)
	{
		char* buff = new char[size];
		m_stringCache.insert(buff);
		return buff;
	}
	void freeStringBuffer(char* buff)
	{
		if (!buff)
			return;
		auto itFinder = m_stringCache.find(buff);
		if (itFinder != m_stringCache.end())
		{
			m_stringCache.erase(itFinder);
			delete[] buff;
		}
	}

	char* moveStringBuffer(const std::string& innerStr)
	{
		char* buff = allocStringBuffer(innerStr.size() + 1);
		memcpy(buff, &innerStr[0], innerStr.size());
		buff[innerStr.size()] = '\0';
		return buff;
	}

protected:
	std::set<char*> m_stringCache;
};

class LastError
{
public:
	LastError()
	{

	}

	virtual void SetLastError(const std::string& lastError)
	{
		m_lastError = lastError;
	}

	virtual const char* GetLastError()
	{
		return m_lastError.c_str();
	}

	virtual void ClearLastError()
	{
		m_lastError.clear();
	}

private:
	std::string m_lastError;
};

class CellImp 
	: public Cell
	, public BufferManager
	, public LastError
{
	friend class WorkSheetImp;

public:
	CellImp(const xlnt::cell& cell)
		: m_cell(cell)
	{

	}
	~CellImp()
	{

	}

	xlnt::cell& raw()
	{
		return m_cell;
	}

public:
	virtual CellType GetValueType() override
	{
		try
		{
			return (CellType)m_cell.data_type();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return CellType::empty_value;
		}
	}

	virtual char* GetStringValue(const char* defVal = "") override
	{
		try
		{
			return moveStringBuffer(Utf8ToAnsi(m_cell.to_string()));
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return const_cast<char*>(defVal);
		}
	}

	virtual bool SetStringValue(const char* strVal) override
	{
		try
		{
			m_cell.value(AnsiToUtf8(strVal));
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetNullValue() override
	{
		try
		{
			m_cell.clear_value();
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetBoolValue(bool boolean_value) override
	{
		try
		{
			m_cell.value(boolean_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool GetBoolValue(bool defVal) override
	{
		try
		{
			return m_cell.value<bool>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SetIntValue(int int_value) override
	{
		try
		{
			m_cell.value(int_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual int GetIntValue(int defVal) override
	{
		try
		{
			return m_cell.value<int>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SetUIntValue(unsigned int int_value) override
	{
		try
		{
			m_cell.value(int_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual unsigned int GetUIntValue(unsigned int defVal) override
	{
		try
		{
			return m_cell.value<unsigned int>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SetLLIntValue(long long int int_value) override
	{
		try
		{
			m_cell.value(int_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual long long int GetLLIntValue(long long int defVal) override
	{
		try
		{
			return m_cell.value<long long int>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SetULLIntValue(unsigned long long int int_value) override
	{
		try
		{
			m_cell.value(int_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual unsigned long long int GetULLIntValue(unsigned long long int defVal) override
	{
		try
		{
			return m_cell.value<unsigned long long int>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SetDoubleValue(double float_value) override
	{
		try
		{
			m_cell.value(float_value);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual double GetDoubleValue(double defVal) override
	{
		try
		{
			return m_cell.value<double>();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return defVal;
		}
	}

	virtual bool SeDateValue(int year_, int month_, int day_) override
	{
		try
		{
			xlnt::date d(year_, month_, day_);
			m_cell.value(d);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool GetDateValue(int& year_, int& month_, int& day_) override
	{
		try
		{
			xlnt::date d = m_cell.value<xlnt::date>();
			year_ = d.year;
			month_ = d.month;
			day_ = d.day;
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetDatetimeValue(int year_, int month_, int day_, int hour_, int minute_, int second_, int microsecond_) override
	{
		try
		{
			xlnt::datetime d(year_, month_, day_, hour_, minute_, second_, microsecond_);
			m_cell.value(d);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool GetDatetimeValue(int& year_, int& month_, int& day_, int& hour_, int& minute_, int& second_, int& microsecond_) override
	{
		try
		{
			xlnt::datetime d = m_cell.value<xlnt::datetime>();
			year_ = d.year;
			month_ = d.month;
			day_ = d.day;
			hour_ = d.hour;
			minute_ = d.minute;
			second_ = d.second;
			microsecond_ = d.microsecond;
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetTimeValue(int hour_, int minute_, int second_, int microsecond_) override
	{
		try
		{
			xlnt::time t(hour_, minute_, second_, microsecond_);
			m_cell.value(t);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool GetTimeValue(int& hour_, int& minute_, int& second_, int& microsecond_) override
	{
		try
		{
			xlnt::time d = m_cell.value<xlnt::time>();
			hour_ = d.hour;
			minute_ = d.minute;
			second_ = d.second;
			microsecond_ = d.microsecond;
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetTimeDeltaValue(int days_, int hours_, int minutes_, int seconds_, int microseconds_) override
	{
		try
		{
			xlnt::timedelta t(days_, hours_, minutes_, seconds_, microseconds_);
			m_cell.value(t);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool GetTimeDeltaValue(int days_, int& hours_, int& minutes_, int& seconds_, int& microseconds_) override
	{
		try
		{
			xlnt::timedelta d = m_cell.value<xlnt::timedelta>();
			days_ = d.days;
			hours_ = d.hours;
			minutes_ = d.minutes;
			seconds_ = d.seconds;
			microseconds_ = d.microseconds;
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool SetFormula(const char* formula) override
	{
		try
		{
			m_cell.formula(formula);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool CopyFrom(Cell* otherCell) override
	{
		try
		{
			CellImp* imp = (CellImp*)otherCell;
			m_cell.value(imp->raw());
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual void FreeString(char* buff) override
	{
		freeStringBuffer(buff);
	}

private:
	xlnt::cell m_cell;
};

class WorkSheetImp
	: public WorkSheet
	, public BufferManager
	, public LastError
{
public:
	WorkSheetImp(xlnt::worksheet& sheet)
		: m_sheet(sheet)
	{

	}

	xlnt::worksheet& raw()
	{
		return m_sheet;
	}

public:
	virtual char* GetSheetTitle() override
	{
		try
		{
			return moveStringBuffer(m_sheet.title());
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual bool SetSheetTitle(const char* title) override
	{
		try
		{
			m_sheet.title(title);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual size_t GetRowCount() override
	{
		try
		{
			xlnt::range range = m_sheet.rows(false);
			return range.length();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual size_t GetColumnCount() override
	{
		try
		{
			xlnt::range range = m_sheet.columns(false);
			return range.length();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual size_t GetNotEmptyRowStart() override
	{
		try
		{
			return m_sheet.lowest_row() - 1;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual size_t GetNotEmptyRowEnd() override
	{
		try
		{
			return m_sheet.highest_row() - 1;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual size_t GetNotEmptyColumnStart() override
	{
		try
		{			
			return m_sheet.lowest_column().index - 1;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual size_t GetNotEmptyColumnEnd() override
	{
		try
		{
			return m_sheet.highest_column().index - 1;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual char* GetRowStringArray(size_t rowIndex, size_t* columnCount) override
	{
		try
		{
			xlnt::range rows = m_sheet.rows(false);
			if (rowIndex >= rows.length())
			{
				columnCount = 0;
				return NULL;
			}
			xlnt::cell_vector row = rows[rowIndex];
			std::vector<std::string> rowData;
			size_t totalSize = 0;
			for (auto cell : row)
			{
				rowData.push_back(Utf8ToAnsi(cell.to_string()));
				totalSize += rowData[rowData.size() - 1].size() + 1;
			}
			char* totalBuff = allocStringBuffer(totalSize);
			size_t start = 0;
			for (size_t i = 0; i < rowData.size(); i++)
			{
				strcpy_s(totalBuff + start, totalSize - start, rowData[i].c_str());
				(totalBuff + start)[rowData[i].size()] = '\0';
				start += rowData[i].size() + 1;
			}
			*columnCount = rowData.size();
			return totalBuff;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			columnCount = 0;
			return NULL;
		}
	}

	virtual Cell* OpenCell(size_t rowIdx, int columnIdx) override
	{
		try
		{
			xlnt::cell cell = m_sheet.cell(xlnt::column_t(columnIdx + 1), xlnt::row_t(rowIdx + 1));
			CellImp* imp = new CellImp(cell);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual Cell* OpenCell(const char* refName) override
	{
		try
		{
			xlnt::cell cell = m_sheet.cell(refName);
			CellImp* imp = new CellImp(cell);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual void CloseCell(Cell* cell) override
	{
		if (!cell)
			return;
		CellImp* imp = (CellImp*)cell;
		delete imp;
	}

	virtual void FreezePanes(Cell* top_left_cell) override
	{
		m_sheet.freeze_panes(((CellImp*)top_left_cell)->raw());
	}

	virtual void UnfreezePanes() override
	{
		m_sheet.unfreeze_panes();
	}

	virtual bool HasFreezePanes() override
	{
		return m_sheet.has_frozen_panes();
	}

	virtual void FreeString(char* buff) override
	{
		freeStringBuffer(buff);
	}

private:
	xlnt::worksheet m_sheet;
};

class WorkBookImp
	: public WorkBook
	, public BufferManager
	, public LastError
{
public:
	~WorkBookImp()
	{
		
	}

public:
	virtual int GetSheetCount() override
	{
		try
		{
			return (int)m_workbook.sheet_count();
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return 0;
		}
	}

	virtual WorkSheet* OpenSheetByIndex(int sheetIndex) override
	{
		try
		{
			xlnt::worksheet sheet = m_workbook.sheet_by_index(sheetIndex);
			WorkSheetImp* imp = new WorkSheetImp(sheet);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual WorkSheet* OpenSheetByTitle(const char* sheetTitle) override
	{
		try
		{
			xlnt::worksheet sheet = m_workbook.sheet_by_title(sheetTitle);
			WorkSheetImp* imp = new WorkSheetImp(sheet);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual bool SetCurrentSheet(WorkSheet* sheet) override
	{
		try
		{
			WorkSheetImp* imp = (WorkSheetImp*)sheet;
			size_t idx = m_workbook.index(imp->raw());
			m_workbook.active_sheet(idx);
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual WorkSheet* OpenCurrentSheet() override
	{
		try
		{
			xlnt::worksheet sheet = m_workbook.active_sheet();
			WorkSheetImp* imp = new WorkSheetImp(sheet);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual WorkSheet* CreateSheet(int atIndex) override
	{
		try
		{
			xlnt::worksheet sheet;
			if (atIndex < 0)
				sheet = m_workbook.create_sheet();
			else
			{
				if (atIndex >= m_workbook.sheet_count())
					return NULL;
				sheet = m_workbook.create_sheet(atIndex);
			}
			WorkSheetImp* imp = new WorkSheetImp(sheet);
			return imp;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual WorkSheet* CloneSheet(WorkSheet* srcSheet, int atIndex) override
	{
		try
		{
			WorkSheetImp* imp = (WorkSheetImp*)srcSheet;

			if (atIndex < 0)
			{
				xlnt::worksheet sheet = m_workbook.copy_sheet(imp->raw());
				WorkSheetImp* newSheet = new WorkSheetImp(sheet);
				return newSheet;
			}
			else
			{
				if (atIndex >= m_workbook.sheet_count())
					return NULL;
				xlnt::worksheet sheet = m_workbook.copy_sheet(imp->raw(), atIndex);
				WorkSheetImp* newSheet = new WorkSheetImp(sheet);
				return newSheet;
			}
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return NULL;
		}
	}

	virtual bool RemoveSheet(WorkSheet* sheet) override
	{
		try
		{
			WorkSheetImp* imp = (WorkSheetImp*)sheet;
			m_workbook.remove_sheet(imp->raw());
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual bool Save(const char* filepath) override
	{
		try
		{
			m_workbook.save(xlnt::path(filepath));
			return true;
		}
		catch (const std::exception& ex)
		{
			SetLastError(ex.what());
			return false;
		}
	}

	virtual void CloseSheet(WorkSheet* sheet) override
	{
		if (!sheet)
			return;
		WorkSheetImp* imp = (WorkSheetImp*)sheet;
		delete imp;
	}

	virtual void FreeString(char* buff) override
	{
		freeStringBuffer(buff);
	}

public:
	bool Open(const char* xlsxFilepath, const char* password)
	{
		std::string sPassword;
		if (password && password[0] != '\0')
		{
			sPassword = password;
		}

		try
		{
			//if the file path exist,load the file to workbook
			if (IsPathExist(xlsxFilepath))
			{
				std::string utf8Str = AnsiToUtf8(xlsxFilepath);

				if (sPassword.empty())
					m_workbook.load(xlnt::path(utf8Str));
				else
					m_workbook.load(xlnt::path(utf8Str), sPassword);
			}
		}
		catch (const std::exception& ex)
		{
			std::string error = ex.what();
			return false;
		}

		return true;
	}

	xlnt::workbook& raw()
	{
		return m_workbook;
	}
	
private:
	xlnt::workbook m_workbook;
};

IKXIAO_API WorkBook* OpenExcel(const char* xlsxFilepath, const char* password)
{
	WorkBookImp* wb = new WorkBookImp();
	if (!wb->Open(xlsxFilepath, password))
	{
		delete wb;
		return NULL;
	}
	return wb;
}

IKXIAO_API void CloseExcel(WorkBook* workbook)
{
	if (!workbook)
		return;
	WorkBookImp* imp = (WorkBookImp*)workbook;
	delete imp;
}
