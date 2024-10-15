#pragma once

#include "globalenv.h"


class CProject;


class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(CString fileName);
	void ReInit();
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
	int countNPD = 0;

private:
	// 초기화 필요한 변수들
	int m_totalProjectNum;
	
	ALL_ACT_TYPE	m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;

	PROJECT* m_AllProjects = NULL;	
	PROJECT m_InterProjects[3] = {0,};  //내부프로젝트
	

	int* m_orderTable[2] = {NULL,NULL};

	Dynamic2DArray m_doingHR;
	Dynamic2DArray m_freeHR;
	

	Dynamic2DArray m_doingTable;
	Dynamic2DArray m_doneTable;
	Dynamic2DArray m_defferTable;
		
	Dynamic2DArray m_incomeTable;
	Dynamic2DArray m_expensesTable;

	
	
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, PROJECT* project);
	void SelectNewProject(int thisWeek);
	

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(PROJECT* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);

	void CCompany::ReadOrder(FILE* fp);
	void ReadProject(FILE* fp);


	void RemoveInterProject(PROJECT* project, int addWeek);
	
}; 

