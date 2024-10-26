#pragma once

#include "globalenv.h"


class CProject;


class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(CString fileName);
	void TableInit();
	void ClearMemory();
		
	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	int CalculateTotalInCome();	
	void PrintCompanyResualt();	
	void write_CompanyInfo(Book* book, Sheet* sheet);
		
	GLOBAL_ENV m_GlobalEnv;
	int m_lastDecisionWeek;
	CString m_XlFileName;
	
	Dynamic2DArray m_totalHR;
	int recruitTerm; // 인원 충감을 계산하는 기간 비율 (100/기간(week) 로 계산)	
	
private:
	// 초기화 필요한 변수들
	int m_totalProjectNum;

	PROJECT* m_AllProjects = NULL;	

	Dynamic2DArray m_orderTable;

	Dynamic2DArray m_doingHR;
	Dynamic2DArray m_freeHR;
	

	Dynamic2DArray m_doingTable;
	Dynamic2DArray m_doneTable;
	Dynamic2DArray m_defferTable;
		
	Dynamic2DArray m_incomeTable;
	Dynamic2DArray m_expensesTable;
	Dynamic2DArray m_balanceTable;

	
	
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, PROJECT* project);
	BOOL IsInternalEnoughHR(int thisWeek, PROJECT* project); // 내부 프로젝트 인원 체크			
	void SelectNewProject(int thisWeek);
	

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(PROJECT* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);
	void RemoveInternalProjectEntry(PROJECT* project, int thisWeek);

	void CCompany::ReadOrder(FILE* fp);
	void ReadProject(FILE* fp);


	void AddInternalProjectEntry(PROJECT* project, int addWeek);
	
}; 

