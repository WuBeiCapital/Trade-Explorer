
/******************************************************************** 
filename:   SQLite.h 
created:    2012-11-05 
author:     firehood 
 
purpose:    SQLite数据库操作类 
*********************************************************************/  
#pragma once  
#include <windows.h>  
#include "sqlite3.h"  

//#pragma comment(lib,"sqlite3.lib")   
//#pragma comment(lib,"SQLite.lib")   
  
typedef BOOL (WINAPI *QueryCallback) (void *para, int n_column, char **column_value, char **column_name);  
  
typedef enum _SQLITE_DATATYPE  
{  
    SQLITE_DATATYPE_INTEGER = SQLITE_INTEGER,  
    SQLITE_DATATYPE_FLOAT  = SQLITE_FLOAT,  
    SQLITE_DATATYPE_TEXT  = SQLITE_TEXT,  
    SQLITE_DATATYPE_BLOB = SQLITE_BLOB,  
    SQLITE_DATATYPE_NULL= SQLITE_NULL,  
}SQLITE_DATATYPE;  
  
class SQLite;  
  
class MITC_BASIC_EXT SQLiteDataReader  
{  
public:  
    SQLiteDataReader(sqlite3_stmt *pStmt);  
    ~SQLiteDataReader();  
public:  
    // 读取一行数据  
    BOOL Read();  
    // 关闭Reader，读取结束后调用  
    void Close();  
    // 总的列数  
    int ColumnCount(void);  
    // 获取某列的名称   
    LPCTSTR GetName(int nCol);  
    // 获取某列的数据类型  
    SQLITE_DATATYPE GetDataType(int nCol);  
    // 获取某列的值(字符串)  
    LPCTSTR GetStringValue(int nCol);  
    // 获取某列的值(整形)  
    int GetIntValue(int nCol);  
    // 获取某列的值(长整形)  
    long GetInt64Value(int nCol);  
    // 获取某列的值(浮点形)  
    double GetFloatValue(int nCol);  
    // 获取某列的值(二进制数据)  
    const BYTE* GetBlobValue(int nCol, int &nLen);  
private:  
    sqlite3_stmt *m_pStmt;  
};  
  
class MITC_BASIC_EXT SQLiteCommand  
{  
public:  
    SQLiteCommand(SQLite* pSqlite);  
    SQLiteCommand(SQLite* pSqlite,LPCTSTR lpSql);  
    ~SQLiteCommand();  
public:  
    // 设置命令  
    BOOL SetCommandText(LPCTSTR lpSql);  
    // 绑定参数（index为要绑定参数的序号，从1开始）  
    BOOL BindParam(int index, LPCTSTR szValue);  
    BOOL BindParam(int index, const int nValue);  
    BOOL BindParam(int index, const double dValue);  
    BOOL BindParam(int index, const unsigned char* blobValue, int nLen);  
    // 执行命令  
    BOOL Excute();  
    // 清除命令（命令不再使用时需调用该接口清除）  
    void Clear();  
private:  
    SQLite *m_pSqlite;  
    sqlite3_stmt *m_pStmt;  
};  
  
class MITC_BASIC_EXT SQLite  
{  
public:  
    SQLite(void);  
    ~SQLite(void);  
public:  
    // 打开数据库  
    BOOL Open(LPCTSTR lpDbFlie);  
    // 关闭数据库  
    void Close();  
  
    // 执行非查询操作（更新或删除）  
    BOOL ExcuteNonQuery(LPCTSTR lpSql);  
    BOOL ExcuteNonQuery(SQLiteCommand* pCmd);  
  
    // 查询  
    SQLiteDataReader ExcuteQuery(LPCTSTR lpSql);  
    // 查询（回调方式）  
    BOOL ExcuteQuery(LPCTSTR lpSql,QueryCallback pCallBack);  
  
    // 开始事务  
    BOOL BeginTransaction();  
    // 提交事务  
    BOOL CommitTransaction();  
    // 回滚事务  
    BOOL RollbackTransaction();  
  
    // 获取上一条错误信息  
    LPCTSTR GetLastErrorMsg();  
public:  
    friend class SQLiteCommand;  
private:  
    sqlite3 *m_db;  
};

//////////////////////////////////////////////////////////////////////////////////
//// CppSQLite3 - A C++ wrapper around the SQLite3 embedded database library.
////
//// Copyright (c) 2004..2007 Rob Groves. All Rights Reserved. rob.groves@btinternet.com
//// 
//// Permission to use, copy, modify, and distribute this software and its
//// documentation for any purpose, without fee, and without a written
//// agreement, is hereby granted, provided that the above copyright notice, 
//// this paragraph and the following two paragraphs appear in all copies, 
//// modifications, and distributions.
////
//// IN NO EVENT SHALL THE AUTHOR BE LIABLE TO ANY PARTY FOR DIRECT,
//// INDIRECT, SPECIAL, INCIDENTAL, OR CONSEQUENTIAL DAMAGES, INCLUDING LOST
//// PROFITS, ARISING OUT OF THE USE OF THIS SOFTWARE AND ITS DOCUMENTATION,
//// EVEN IF THE AUTHOR HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
////
//// THE AUTHOR SPECIFICALLY DISCLAIMS ANY WARRANTIES, INCLUDING, BUT NOT
//// LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
//// PARTICULAR PURPOSE. THE SOFTWARE AND ACCOMPANYING DOCUMENTATION, IF
//// ANY, PROVIDED HEREUNDER IS PROVIDED "AS IS". THE AUTHOR HAS NO OBLIGATION
//// TO PROVIDE MAINTENANCE, SUPPORT, UPDATES, ENHANCEMENTS, OR MODIFICATIONS.
////
//// V3.0		03/08/2004	-Initial Version for sqlite3
////
//// V3.1		16/09/2004	-Implemented getXXXXField using sqlite3 functions
////						-Added CppSQLiteDB3::tableExists()
////
//// V3.2		01/07/2005	-Fixed execScalar to handle a NULL result
////			12/07/2007	-Added CppSQLiteDB::IsAutoCommitOn()
////						-Added int64 functions to CppSQLite3Query
////						-Added Name based parameter binding to CppSQLite3Statement.
//////////////////////////////////////////////////////////////////////////////////
//#ifndef _CppSQLite3_H_
//#define _CppSQLite3_H_
//
//#include "sqlite3.h"
//#include <cstdio>
//#include <cstring>
//
//#define CPPSQLITE_ERROR 1000
//
//class CppSQLite3Exception
//{
//public:
//
//    CppSQLite3Exception(const int nErrCode,
//                    char* szErrMess,
//                    bool bDeleteMsg=true);
//
//    CppSQLite3Exception(const CppSQLite3Exception&  e);
//
//    virtual ~CppSQLite3Exception();
//
//    const int errorCode() { return mnErrCode; }
//
//    const char* errorMessage() { return mpszErrMess; }
//
//    static const char* errorCodeAsString(int nErrCode);
//
//private:
//
//    int mnErrCode;
//    char* mpszErrMess;
//};
//
//
//class CppSQLite3Buffer
//{
//public:
//
//    CppSQLite3Buffer();
//
//    ~CppSQLite3Buffer();
//
//    const char* format(const char* szFormat, ...);
//
//    operator const char*() { return mpBuf; }
//
//    void clear();
//
//private:
//
//    char* mpBuf;
//};
//
//
//class CppSQLite3Binary
//{
//public:
//
//    CppSQLite3Binary();
//
//    ~CppSQLite3Binary();
//
//    void setBinary(const unsigned char* pBuf, int nLen);
//    void setEncoded(const unsigned char* pBuf);
//
//    const unsigned char* getEncoded();
//    const unsigned char* getBinary();
//
//    int getBinaryLength();
//
//    unsigned char* allocBuffer(int nLen);
//
//    void clear();
//
//private:
//
//    unsigned char* mpBuf;
//    int mnBinaryLen;
//    int mnBufferLen;
//    int mnEncodedLen;
//    bool mbEncoded;
//};
//
//
//class CppSQLite3Query
//{
//public:
//
//    CppSQLite3Query();
//
//    CppSQLite3Query(const CppSQLite3Query& rQuery);
//
//    CppSQLite3Query(sqlite3* pDB,
//				sqlite3_stmt* pVM,
//                bool bEof,
//                bool bOwnVM=true);
//
//    CppSQLite3Query& operator=(const CppSQLite3Query& rQuery);
//
//    virtual ~CppSQLite3Query();
//
//    int numFields();
//
//    int fieldIndex(const char* szField);
//    const char* fieldName(int nCol);
//
//    const char* fieldDeclType(int nCol);
//    int fieldDataType(int nCol);
//
//    const char* fieldValue(int nField);
//    const char* fieldValue(const char* szField);
//
//    int getIntField(int nField, int nNullValue=0);
//    int getIntField(const char* szField, int nNullValue=0);
//
//	sqlite_int64 getInt64Field(int nField, sqlite_int64 nNullValue=0);
//	sqlite_int64 getInt64Field(const char* szField, sqlite_int64 nNullValue=0);
//
//    double getFloatField(int nField, double fNullValue=0.0);
//    double getFloatField(const char* szField, double fNullValue=0.0);
//
//    const char* getStringField(int nField, const char* szNullValue="");
//    const char* getStringField(const char* szField, const char* szNullValue="");
//
//    const unsigned char* getBlobField(int nField, int& nLen);
//    const unsigned char* getBlobField(const char* szField, int& nLen);
//
//	bool fieldIsNull(int nField);
//    bool fieldIsNull(const char* szField);
//
//    bool eof();
//
//    void nextRow();
//
//    void finalize();
//
//private:
//
//    void checkVM();
//
//	sqlite3* mpDB;
//    sqlite3_stmt* mpVM;
//    bool mbEof;
//    int mnCols;
//    bool mbOwnVM;
//};
//
//
//class CppSQLite3Table
//{
//public:
//
//    CppSQLite3Table();
//
//    CppSQLite3Table(const CppSQLite3Table& rTable);
//
//    CppSQLite3Table(char** paszResults, int nRows, int nCols);
//
//    virtual ~CppSQLite3Table();
//
//    CppSQLite3Table& operator=(const CppSQLite3Table& rTable);
//
//    int numFields();
//
//    int numRows();
//
//    const char* fieldName(int nCol);
//
//    const char* fieldValue(int nField);
//    const char* fieldValue(const char* szField);
//
//    int getIntField(int nField, int nNullValue=0);
//    int getIntField(const char* szField, int nNullValue=0);
//
//    double getFloatField(int nField, double fNullValue=0.0);
//    double getFloatField(const char* szField, double fNullValue=0.0);
//
//    const char* getStringField(int nField, const char* szNullValue="");
//    const char* getStringField(const char* szField, const char* szNullValue="");
//
//    bool fieldIsNull(int nField);
//    bool fieldIsNull(const char* szField);
//
//    void setRow(int nRow);
//
//    void finalize();
//
//private:
//
//    void checkResults();
//
//    int mnCols;
//    int mnRows;
//    int mnCurrentRow;
//    char** mpaszResults;
//};
//
//
//class CppSQLite3Statement
//{
//public:
//
//    CppSQLite3Statement();
//
//    CppSQLite3Statement(const CppSQLite3Statement& rStatement);
//
//    CppSQLite3Statement(sqlite3* pDB, sqlite3_stmt* pVM);
//
//    virtual ~CppSQLite3Statement();
//
//    CppSQLite3Statement& operator=(const CppSQLite3Statement& rStatement);
//
//    int execDML();
//
//    CppSQLite3Query execQuery();
//
//    void bind(int nParam, const char* szValue);
//    void bind(int nParam, const int nValue);
//    void bind(int nParam, const double dwValue);
//    void bind(int nParam, const unsigned char* blobValue, int nLen);
//    void bindNull(int nParam);
//
//	int bindParameterIndex(const char* szParam);
//    void bind(const char* szParam, const char* szValue);
//    void bind(const char* szParam, const int nValue);
//    void bind(const char* szParam, const double dwValue);
//    void bind(const char* szParam, const unsigned char* blobValue, int nLen);
//    void bindNull(const char* szParam);
//
//	void reset();
//
//    void finalize();
//
//private:
//
//    void checkDB();
//    void checkVM();
//
//    sqlite3* mpDB;
//    sqlite3_stmt* mpVM;
//};
//
//
//class CppSQLite3DB
//{
//public:
//
//    CppSQLite3DB();
//
//    virtual ~CppSQLite3DB();
//
//    void open(const char* szFile);
//
//    void close();
//
//	bool tableExists(const char* szTable);
//
//    int execDML(const char* szSQL);
//
//    CppSQLite3Query execQuery(const char* szSQL);
//
//    int execScalar(const char* szSQL, int nNullValue=0);
//
//    CppSQLite3Table getTable(const char* szSQL);
//
//    CppSQLite3Statement compileStatement(const char* szSQL);
//
//    sqlite_int64 lastRowId();
//
//    void interrupt() { sqlite3_interrupt(mpDB); }
//
//    void setBusyTimeout(int nMillisecs);
//
//    static const char* SQLiteVersion() { return SQLITE_VERSION; }
//    static const char* SQLiteHeaderVersion() { return SQLITE_VERSION; }
//    static const char* SQLiteLibraryVersion() { return sqlite3_libversion(); }
//    static int SQLiteLibraryVersionNumber() { return sqlite3_libversion_number(); }
//
//	bool IsAutoCommitOn();
//
//private:
//
//    CppSQLite3DB(const CppSQLite3DB& db);
//    CppSQLite3DB& operator=(const CppSQLite3DB& db);
//
//    sqlite3_stmt* compile(const char* szSQL);
//
//    void checkDB();
//
//    sqlite3* mpDB;
//    int mnBusyTimeoutMs;
//};
//
//#endif

#if !defined(_X64)
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/Win32/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/Win32/sqlite3.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/Win32/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/Win32/sqlite3.lib"
		#endif
	#endif
#else
	#if defined(_DEBUG)
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Debug/x64/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Debug/x64/sqlite3.lib"
		#endif
	#else
		#if defined(_AFXEXT)
			#define AUTOLIBNAME "../../Lib/Release/x64/sqlite3.lib"
		#else
			#define AUTOLIBNAME "../../Lib/Release/x64/sqlite3.lib"
		#endif
	#endif
#endif

// Perform autolink here:
#pragma message( "automatically link with (" AUTOLIBNAME ")")
#pragma comment(lib, AUTOLIBNAME)