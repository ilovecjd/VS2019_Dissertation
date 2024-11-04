#pragma once

#include "GlobalEnv.h"

#include <xlnt/xlnt.hpp>
#include <filesystem>
#include <string>

class CCreator
{
public:
	CCreator();
	CCreator(int count);
	~CCreator();

	// song
public :
	
	int m_totalProjectNum;	
	Dynamic2DArray m_orderTable;
	
	BOOL Init(GLOBAL_ENV* pGlobalEnv, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern);	
	void Save(CString filename, CString strInSheetName);
		
private:	
	GLOBAL_ENV m_GlobalEnv;
	ALL_ACT_TYPE m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;
	DynamicProjectArray m_pProjects;

	int CreateOrderTable();
	int CreateProjects();
	BOOL CreateActivities(PROJECT* pProject);
	int CalculateHRAndProfit(PROJECT* pProject);
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule(PROJECT* pProject);

	void WriteProjet(FILE* fp);

	int CreateAllProjects();
	int CreateInternalProject(int Id, int week);
	int CraterExternalProject(int Id, int week);
	BOOL CreateInternalActivities(PROJECT* pProject);

public:
	void PrintProjectInfo();
};