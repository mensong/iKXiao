#include "pch.h"
#include "iKXiao.h"
#include <xlnt/xlnt.hpp>
#include <io.h>
#include <set>
#include <vector>

class iKXiaoImp
{
public:
	iKXiaoImp()
	{

	}
	~iKXiaoImp()
	{
		for (auto it = m_stringCache.begin(); it != m_stringCache.end(); ++it)
		{
			delete[] (*it);
		}
		m_stringCache.clear();
	}

	IK_WORKBOOK Open(const char* xlsxFilepath, const char* password)
	{
		//set data
		m_filepath = xlsxFilepath;
		if (password && password[0] != '\0')
		{
			m_password = password;
		}

		try
		{
			//if the file path exist,load the file to workbook
			if (IsPathExist(xlsxFilepath))
			{
				if (m_password.empty())
					m_workbook.load(xlnt::path(xlsxFilepath));
				else
					m_workbook.load(xlnt::path(xlsxFilepath), m_password);
			}
		}
		catch (const std::exception&)
		{
			return NULL;
		}

		return this;
	}

	int GetSheetCount()
	{
		return (int)m_workbook.sheet_count();
	}

	int GetSheetIndexByTitle(const char* sheetTitle)
	{
		auto titles = m_workbook.sheet_titles();
		for (size_t i = 0; i < titles.size(); i++)
		{
			if (titles[i] == sheetTitle)
				return i;
		}
		return -1;
	}

	char* GetSheetTitle(int sheetIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return moveStringBuffer(sheet.title());
		}
		catch (const std::exception&)
		{
			return NULL;
		}
	}

	bool SetCurrentSheet(int sheetIndex)
	{
		if (sheetIndex < 0 && sheetIndex >= GetSheetCount())
			return false;

		try
		{
			m_workbook.active_sheet(sheetIndex);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}		
	}

	//atIndex==-1时在最后创建
	int CreateSheet(int atIndex = -1)
	{
		try
		{
			if (atIndex < 0)
			{
				m_workbook.create_sheet();
				return GetSheetCount() - 1;
			}
			else
			{
				m_workbook.create_sheet((size_t)atIndex);
				return atIndex;
			}
		}
		catch (const std::exception&)
		{
			return -1;
		}
	}

	int CopySheet(int srcIndex, int atIndex = -1)
	{
		try
		{
			auto sheet = getSheetByIndex(srcIndex);

			if (atIndex < 0)
			{
				m_workbook.copy_sheet(sheet);
				return GetSheetCount() - 1;
			}
			else
			{
				m_workbook.copy_sheet(sheet, (size_t)atIndex);
				return atIndex;
			}
		}
		catch (const std::exception&)
		{
			return -1;
		}
	}

	bool RemoveSheet(int sheetIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			m_workbook.remove_sheet(sheet);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetSheetTitle(int sheetIndex, const char* title)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.title(title);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	size_t GetRowCount(int sheetIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			auto rows = sheet.rows(false);
			return rows.length();
		}
		catch (const std::exception&)
		{
			return 0;
		}
	}

	char* GetRowStringArray(int sheetIndex, size_t rowIndex, size_t* columnCount)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			auto rows = sheet.rows(false);
			if (rowIndex >= rows.length())
			{
				columnCount = 0;
				return NULL;
			}
			auto row = rows[rowIndex];
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
		catch (const std::exception&)
		{
			columnCount = 0;
			return NULL;
		}
	}

	CellType GetCellType(int sheetIndex, size_t rowIndex, size_t columnIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			auto cell = sheet.cell(xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1));
			return (CellType)cell.data_type();
		}
		catch (const std::exception&)
		{
			return CellType::empty_value;
		}
	}

	char* GetCellStringValue(int sheetIndex, size_t rowIndex, size_t columnIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			auto cell = sheet.cell(xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1));
			return moveStringBuffer(Utf8ToAnsi(cell.to_string()));
		}
		catch (const std::exception&)
		{
			return NULL;
		}
	}

	char* GetCellStringValueByRefName(int sheetIndex, const char* refName)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);			
			auto cell = sheet.cell(refName);
			return moveStringBuffer(Utf8ToAnsi(cell.to_string()));
		}
		catch (const std::exception&)
		{
			return NULL;
		}
	}

	bool SetCellStringValue(int sheetIndex, size_t rowIndex, size_t columnIndex, const char* strVal)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(AnsiToUtf8(strVal));
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellNullValue(int sheetIndex, size_t rowIndex, size_t columnIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).clear_value();
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellBoolValue(int sheetIndex, size_t rowIndex, size_t columnIndex, bool boolean_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(boolean_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool GetCellBoolValue(int sheetIndex, size_t rowIndex, size_t columnIndex, bool defVal = false)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<bool>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, int int_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(int_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	int GetCellIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, int defVal = 0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<int>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellUIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int int_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(int_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	unsigned int GetCellUIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int defVal = 0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<unsigned int>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellLLIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, long long int int_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(int_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	long long int GetCellLLIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, long long int defVal = 0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<long long int>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellULLIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int int_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(int_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	unsigned long long int GetCellULLIntValue(int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int defVal = 0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<unsigned long long int>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellDoubleValue(int sheetIndex, size_t rowIndex, size_t columnIndex, double float_value)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(float_value);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	double GetCellDoubleValue(int sheetIndex, size_t rowIndex, size_t columnIndex, double defVal = 0.0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			return sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<double>();
		}
		catch (const std::exception&)
		{
			return defVal;
		}
	}

	bool SetCellDateValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int year_, int month_, int day_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::date d(year_, month_, day_);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(d);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool GetCellDateValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int& year_, int& month_, int& day_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::date d = sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<xlnt::date>();
			year_ = d.year;
			month_ = d.month;
			day_ = d.day;
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellDatetimeValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::datetime d(year_, month_, day_, hour_, minute_, second_, microsecond_);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(d);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool GetCellDatetimeValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int& year_, int& month_, int& day_, int& hour_, int& minute_, int& second_, int& microsecond_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::datetime d = sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<xlnt::datetime>();
			year_ = d.year;
			month_ = d.month;
			day_ = d.day;
			hour_ = d.hour;
			minute_ = d.minute;
			second_ = d.second;
			microsecond_ = d.microsecond;
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellTimeValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0)
	{		
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::time t(hour_, minute_, second_, microsecond_);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(t);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool GetCellTimeValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int& hour_, int& minute_, int& second_, int& microsecond_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::time d = sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<xlnt::time>();
			hour_ = d.hour;
			minute_ = d.minute;
			second_ = d.second;
			microsecond_ = d.microsecond;
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellTimeDeltaValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int days_, int hours_, int minutes_, int seconds_, int microseconds_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::timedelta t(days_, hours_, minutes_, seconds_, microseconds_);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value(t);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool GetCellTimeDeltaValue(int sheetIndex, size_t rowIndex, size_t columnIndex,
		int days_, int& hours_, int& minutes_, int& seconds_, int& microseconds_)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::timedelta d = sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).value<xlnt::timedelta>();
			days_ = d.days;
			hours_ = d.hours;
			minutes_ = d.minutes;
			seconds_ = d.seconds;
			microseconds_ = d.microseconds;
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool SetCellFormula(int sheetIndex, size_t rowIndex, size_t columnIndex, const char* formula)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			sheet.cell(
				xlnt::column_t(columnIndex + 1), xlnt::row_t(rowIndex + 1)
			).formula(formula);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool CopyCell(int sheetIndex, size_t fromRowIndex, size_t fromColumnIndex, size_t toRowIndex, size_t toColumnIndex)
	{
		try
		{
			auto sheet = getSheetByIndex(sheetIndex);
			xlnt::cell c = sheet.cell(
				xlnt::column_t(fromColumnIndex + 1), xlnt::row_t(fromRowIndex + 1));
			sheet.cell(
				xlnt::column_t(toColumnIndex + 1), xlnt::row_t(toRowIndex + 1)
			).value(c);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

	bool Save(const char* filepath, const char* password)
	{
		try
		{
			if (!password || password[0] == '\0')
				m_workbook.save(xlnt::path(filepath));
			else
				m_workbook.save(xlnt::path(filepath), password);
			return true;
		}
		catch (const std::exception&)
		{
			return false;
		}
	}

public:
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
		m_stringCache.erase(buff);
		delete[] buff;		
	}

	char* moveStringBuffer(const std::string& innerStr)
	{
		char* buff = allocStringBuffer(innerStr.size() + 1);
		memcpy(buff, &innerStr[0], innerStr.size());
		buff[innerStr.size()] = '\0';
		return buff;
	}

	xlnt::worksheet getSheetByIndex(int sheetIndex)
	{
		try
		{
			switch (sheetIndex)
			{
			case -1:
				return m_workbook.active_sheet();
			case -2:
				return m_workbook.sheet_by_index(m_workbook.sheet_count() - 1);
			default:
				return m_workbook.sheet_by_index(sheetIndex);
				break;
			}
		}
		catch (const std::exception&)
		{
			return m_workbook.active_sheet();
		}
	}

private:	
	std::set<char*> m_stringCache;

private:
	xlnt::workbook m_workbook;
	std::string m_filepath;
	std::string m_password;
};

IKXIAO_API IK_WORKBOOK OpenExcel(const char* xlsxFilepath, const char* password)
{
	iKXiaoImp* imp = new iKXiaoImp();
	IK_WORKBOOK ret = imp->Open(xlsxFilepath, password);
	if (!ret)
		delete imp;
	return ret;
}

IKXIAO_API int GetSheetCount(IK_WORKBOOK workbook)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetSheetCount();
}

IKXIAO_API int GetSheetIndexByTitle(IK_WORKBOOK workbook, const char* sheetTitle)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetSheetIndexByTitle(sheetTitle);
}

IKXIAO_API char* GetSheetTitle(IK_WORKBOOK workbook, int sheetIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetSheetTitle(sheetIndex);
}

IKXIAO_API bool SetCurrentSheet(IK_WORKBOOK workbook, int sheetIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCurrentSheet(sheetIndex);
}

IKXIAO_API int CreateSheet(IK_WORKBOOK workbook, int atIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->CreateSheet(atIndex);
}

IKXIAO_API int CopySheet(IK_WORKBOOK workbook, int srcIndex, int atIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->CopySheet(srcIndex, atIndex);
}

IKXIAO_API bool RemoveSheet(IK_WORKBOOK workbook, int sheetIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->RemoveSheet(sheetIndex);
}

IKXIAO_API bool SetSheetTitle(IK_WORKBOOK workbook, int sheetIndex, const char* title)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetSheetTitle(sheetIndex, title);
}

IKXIAO_API size_t GetRowCount(IK_WORKBOOK workbook, int sheetIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetRowCount(sheetIndex);
}

IKXIAO_API char* GetRowStringArray(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t* columnCount)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetRowStringArray(sheetIndex, rowIndex, columnCount);
}

IKXIAO_API CellType GetCellType(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellType(sheetIndex, rowIndex, columnIndex);
}

IKXIAO_API char* GetCellStringValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellStringValue(sheetIndex, rowIndex, columnIndex);
}

IKXIAO_API char* GetCellStringValueByRefName(IK_WORKBOOK workbook, int sheetIndex, const char* refName)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellStringValueByRefName(sheetIndex, refName);
}

IKXIAO_API bool SetCellStringValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, const char* strVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellStringValue(sheetIndex, rowIndex, columnIndex, strVal);
}

IKXIAO_API bool SetCellNullValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellNullValue(sheetIndex, rowIndex, columnIndex);
}

IKXIAO_API bool SetCellBoolValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, bool boolean_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellBoolValue(sheetIndex, rowIndex, columnIndex, boolean_value);
}

IKXIAO_API bool GetCellBoolValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, bool defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellBoolValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int int_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellIntValue(sheetIndex, rowIndex, columnIndex, int_value);
}

IKXIAO_API int GetCellIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellIntValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellUIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int int_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellUIntValue(sheetIndex, rowIndex, columnIndex, int_value);
}

IKXIAO_API unsigned int GetCellUIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned int defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellUIntValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellLLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, long long int int_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellLLIntValue(sheetIndex, rowIndex, columnIndex, int_value);
}

IKXIAO_API long long int GetCellLLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, long long int defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellLLIntValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellULLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int int_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellULLIntValue(sheetIndex, rowIndex, columnIndex, int_value);
}

IKXIAO_API unsigned long long int GetCellULLIntValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, unsigned long long int defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellULLIntValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellDoubleValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, double float_value)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellDoubleValue(sheetIndex, rowIndex, columnIndex, float_value);
}

IKXIAO_API double GetCellDoubleValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, double defVal)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellDoubleValue(sheetIndex, rowIndex, columnIndex, defVal);
}

IKXIAO_API bool SetCellDateValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int year_, int month_, int day_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellDateValue(sheetIndex, rowIndex, columnIndex, year_, month_, day_);
}

IKXIAO_API bool GetCellDateValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& year_, int& month_, int& day_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellDateValue(sheetIndex, rowIndex, columnIndex, year_, month_, day_);
}

IKXIAO_API bool SetCellDatetimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int year_, int month_, int day_, int hour_, int minute_, int second_, int microsecond_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellDatetimeValue(sheetIndex, rowIndex, columnIndex, year_, month_, day_, hour_, minute_, second_, microsecond_);
}

IKXIAO_API bool GetCellDatetimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& year_, int& month_, int& day_, int& hour_, int& minute_, int& second_, int& microsecond_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellDatetimeValue(sheetIndex, rowIndex, columnIndex, year_, month_, day_, hour_, minute_, second_, microsecond_);
}

IKXIAO_API bool SetCellTimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int hour_, int minute_, int second_, int microsecond_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellTimeValue(sheetIndex, rowIndex, columnIndex, hour_, minute_, second_, microsecond_);
}

IKXIAO_API bool GetCellTimeValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int& hour_, int& minute_, int& second_, int& microsecond_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellTimeValue(sheetIndex, rowIndex, columnIndex, hour_, minute_, second_, microsecond_);
}

IKXIAO_API bool SetCellTimeDeltaValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int days_, int hours_, int minutes_, int seconds_, int microseconds_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellTimeDeltaValue(sheetIndex, rowIndex, columnIndex, days_, hours_, minutes_, seconds_, microseconds_);
}

IKXIAO_API bool GetCellTimeDeltaValue(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, int days_, int& hours_, int& minutes_, int& seconds_, int& microseconds_)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->GetCellTimeDeltaValue(sheetIndex, rowIndex, columnIndex, days_, hours_, minutes_, seconds_, microseconds_);
}

IKXIAO_API bool SetCellFormula(IK_WORKBOOK workbook, int sheetIndex, size_t rowIndex, size_t columnIndex, const char* formula)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->SetCellFormula(sheetIndex, rowIndex, columnIndex, formula);
}

IKXIAO_API bool CopyCell(IK_WORKBOOK workbook, int sheetIndex, size_t fromRowIndex, size_t fromColumnIndex, size_t toRowIndex, size_t toColumnIndex)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->CopyCell(sheetIndex, fromRowIndex, fromColumnIndex, toRowIndex, toColumnIndex);
}

IKXIAO_API bool SaveExcel(IK_WORKBOOK workbook, const char* filepath, const char* password)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	return imp->Save(filepath, password);
}

IKXIAO_API void FreeString(IK_WORKBOOK workbook, char* buff)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	imp->freeStringBuffer(buff);
}

IKXIAO_API void CloseExcel(IK_WORKBOOK workbook)
{
	iKXiaoImp* imp = (iKXiaoImp*)workbook;
	delete imp;
}
