#include "pch.h"
#include "GlobalEnv.h"
#include "Creator.h"
#include <cctype>   // toupper 함수를 사용하기 위해 필요
#include "hidden_libxl.h"



CCreator::CCreator()
{
	
}
CCreator::CCreator(int count)
{
	countNPD = count;
}

CCreator::~CCreator()
{	
}

// 문제를 생성한다. 글로벌 환경에서 발생 할 수 있는 프로젝트들을 발생 시키고 파일로 저장한다.
//BOOL CCreator::Init(int type, int ID, int ODate, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern)
BOOL CCreator::Init(GLOBAL_ENV* pGlobalEnv, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern)
{	
	*(&m_GlobalEnv) = *pGlobalEnv;
	*(&m_ActType) = *pActType;
	*(&m_ActPattern) = *pActPattern;
	m_pProjects.Resize(0,10);

	m_totalProjectNum = CreateAllProjects();

	// 디버깅 정보를 출력한다.
	//PrintProjectInfo();

	//CreateOrderTable();//m_totalProjectNum 생성 (내부프로젝트 3개 만큼 크게)
	//m_pProjects = new PROJECT[m_totalProjectNum];
	//CreateProjects();

	return TRUE;
}


/*
1) 처음은 내부 프로젝트로

2) 이번주 발생 프로젝트 갯수
프로젝트 수 만큼 발생

*/
int CCreator::CreateAllProjects()
{
	int prjectId = 0;
	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		if (0 == week % (52)) // 1년에 하나씩 내부 프로젝트 발생
		{
			CreateInternalProject(prjectId++,week);
		}

		int newCnt = PoissonRandom(m_GlobalEnv.WeeklyProb);	// 이번주 발생하는 프로젝트 갯수

		for (int i = 0; i < newCnt; i++) // 외부 프로젝트 발생
		{
			CraterExternalProject(prjectId++,week);

		}
	}

	return prjectId;
}

int CCreator::CreateInternalProject(int Id,int week)
{
	PROJECT Project;
	
	memset(&Project, 0, sizeof(struct PROJECT));

	Project.category = 1;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	Project.ID = Id;		// 프로젝트의 번호	
	Project.orderDate = week;	// 발주일
	Project.startAvail = week;  // 시작 가능일. 내부는 바로 진행가능
	Project.isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	Project.experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
	Project.winProb = 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
	Project.nCashFlows = MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

	CreateInternalActivities(&Project);					//m_activities[MAX_ACT] 계산
	Project.profit = CalculateHRAndProfit(&Project); // 총 수익을 계산한다.
	CalculatePaymentSchedule(&Project);			//m_cashFlows[MAX_N_CF] 계산		


	m_pProjects[0][Id] = Project;

	return 0;
}
// 내부 프로젝트의 활동들을 생성한다.
BOOL CCreator::CreateInternalActivities(PROJECT* pProject) {
	//song 사용하지 않는 멤버 변수와 지역 변수들 삭제 하자	
	int i;
	int probability;
	int Lb = 0;
	int UB = 0;
	int maxLoop;
	int totalDuration;
	int tempDuration;

	probability = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
	

	////////////////////////////////////////////
	// 프로젝트 타입관련 정보	
	pProject->projectType = 4;

	//프로젝트 패턴 관련 정보
	pProject->activityPattern = 5;

	Lb = m_ActType.asIntArray[pProject->projectType][2];	// 엑셀 4열의 "최소기간"
	UB = m_ActType.asIntArray[pProject->projectType][3];	// 엑셀 5열의 "최대기간"
	
	totalDuration = RandomBetween(Lb, UB);
	pProject->duration = totalDuration;
	pProject->endDate = pProject->startAvail + totalDuration - 1;// song??


	Lb = 0;
	UB = 0;
	maxLoop = m_ActPattern.asIntArray[pProject->activityPattern][0];//활동수 !!! -1 에 주의
	pProject->numActivities = maxLoop;

	// 활동 생성
	for (i = 0; i < maxLoop; ++i) {
		Lb += m_ActPattern.asIntArray[pProject->activityPattern][1 + ((i) * 5)];// !!! -1 에 주의
		UB += m_ActPattern.asIntArray[pProject->activityPattern][2 + ((i) * 5)];// !!! -1 에 주의
		probability = RandomBetween(Lb, UB);
		tempDuration = totalDuration * probability / 100;

		if (tempDuration == 0) {
			tempDuration = 1;
		}

		if (i == 0) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail;
			pProject->activities[i].endDate = pProject->startAvail - 1 + tempDuration;
		}
		else if (i == 1) {
			pProject->activities[i].duration = totalDuration - pProject->activities[0].duration;
			pProject->activities[i].startDate = pProject->activities[0].endDate + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else if (i == 2) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail - 1 + totalDuration - tempDuration + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->activities[2].startDate - tempDuration;
			pProject->activities[i].endDate = pProject->activities[2].startDate - 1;
		}
	}
	return TRUE;
}

int CCreator::CraterExternalProject(int Id, int week)
{
	PROJECT Project;

	memset(&Project, 0, sizeof(struct PROJECT));

	Project.category = 0;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	Project.ID = Id;		// 프로젝트의 번호	
	Project.orderDate = week;	// 발주일
	Project.startAvail = week + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
	Project.isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	Project.experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
	Project.winProb = 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
	Project.nCashFlows = MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

	CreateActivities(&Project);					//m_activities[MAX_ACT] 계산
	Project.profit = CalculateHRAndProfit(&Project); // 총 수익을 계산한다.
	CalculatePaymentSchedule(&Project);			//m_cashFlows[MAX_N_CF] 계산		


	m_pProjects[0][Id] = Project;
		
	return 0;

}

/////////////////////////////////////////////////////////////////////////
// 프로젝트 발주(발생) 현황 생성, 프로젝트는 최대 크기 만큼 설정한다.
// 시뮬레이션보다 길게 작성한다.
int CCreator::CreateOrderTable()
{
	int cnt = 0, sum = 0;

	m_orderTable.Resize(2, m_GlobalEnv.maxWeek);

	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		cnt = PoissonRandom(m_GlobalEnv.WeeklyProb);	// 이번주 발생하는 프로젝트 갯수		
		m_orderTable[ORDER_SUM][week] = sum;			// 누계
		m_orderTable[ORDER_ORD][week] = cnt;			// 발생 프로젝트갯수
		sum = sum + cnt;	// 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다.
	}
	//m_OutProjectNum = sum;
	m_totalProjectNum = sum + countNPD;// +3;// 생성될 내부프로젝트 최대 갯수만큼 더한다.
	return 0;
}

int CCreator::CreateProjects()
{
	int projectId = 0;
	int startNum = 0;
	int endNum = 0;
	int preTotal = 0;

	// 외부 프로젝트 생성
	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		//preTotal = m_orderTable[ORDER_SUM][week];			// 지난주까지의 발주 프로젝트 누계
		//startNum = preTotal + 1;						// 신규프로젝트이 시작번호 = 누계 +1
		//endNum = preTotal + m_orderTable[ORDER_ORD][week];	// 마지막 프로젝트의 시작번호 = 지난주 누계 + 이번주 발생건수

		//if ((startNum != 0) && (startNum <= endNum))
		//{
		//	for (projectId = startNum; projectId <= endNum; projectId++)
		//	{
		//		PROJECT* pProject;				
		//		pProject =  &m_pProjects[projectId-1];
		//		memset(pProject, 0, sizeof(struct PROJECT));

		//		pProject->category = 0;		// 프로젝트 분류 (0: 외부 / 1: 내부)
		//		pProject->ID = projectId;		// 프로젝트의 번호	
		//		pProject->orderDate = week;	// 발주일
		//		pProject->startAvail = week  + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
		//		pProject->isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
		//		pProject->experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
		//		pProject->winProb = 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
		//		pProject->nCashFlows = MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

		//		CreateActivities(pProject);					//m_activities[MAX_ACT] 계산
		//		pProject->profit = CalculateHRAndProfit(pProject); // 총 수익을 계산한다.
		//		CalculatePaymentSchedule(pProject);			//m_cashFlows[MAX_N_CF] 계산				
		//	}
		//}
	}

	// 내부 프로젝트 생성
	for (int i = 0; i < countNPD; i++)
	{	
		//PROJECT* pProject;
		//int duration = 40+i*4;  // 10개월, 11개월, 12개월
		//int startDate = i * 4 * 12;

		//projectId = m_OutProjectNum + 1 + i;
		//pProject = &m_pProjects[projectId-1];
		//memset(pProject, 0, sizeof(struct PROJECT));

		//pProject->category = 1;		// 프로젝트 분류 (0: 외부 / 1: 내부)
		//pProject->ID = projectId;		// 프로젝트의 번호	
		//pProject->orderDate = startDate;	// 발주일
		//pProject->startAvail = startDate; // 바로시작 가능
		//pProject->isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
		//pProject->experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
		//pProject->winProb = 30 + i*10;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 
		//pProject->nCashFlows = 0;			// 비용 지급 횟수(규모에 따라 변경 가능)
		//
		//pProject->endDate = startDate + duration - 1;		// 프로젝트 종료일
		//pProject->duration = duration;		// 프로젝트의 총 기간

		////CalculatePaymentSchedule(pProject);			//m_cashFlows[MAX_N_CF] 계산	
		//pProject->profit = m_GlobalEnv.Cash_Init /12/4;	// 주당 기대수익으로 변경
		//													
		////1CreateActivities(pProject);			//m_activities[MAX_ACT] 계산 // 활동
		//pProject->numActivities = 1;          // 총 활동 수
		//pProject->activities[0].activityType = 1;// 활동에 관한 정보를 기록하는 배열	
		//pProject->activities[0].duration = duration;      // 활동 기간
		//pProject->activities[0].startDate = startDate;     // 시작 날짜
		//pProject->activities[0].endDate = startDate + duration - 1;       // 종료 날짜
		//pProject->activities[0].highSkill = m_GlobalEnv.Hr_Init_H / 2;     // 높은 기술 수준 인력 수
		//pProject->activities[0].midSkill = m_GlobalEnv.Hr_Init_H / 2;      // 중간 기술 수준 인력 수
		//pProject->activities[0].lowSkill = m_GlobalEnv.Hr_Init_H / 2;      // 낮은 기술 수준 인력 수

	}

	return 0;

}


BOOL CCreator::CreateActivities(PROJECT* pProject) {
	//song 사용하지 않는 멤버 변수와 지역 변수들 삭제 하자	
	int i;
	int probability;
	int Lb = 0;
	int UB = 0;
	int maxLoop;
	int totalDuration;
	int tempDuration;

	probability = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
	maxLoop = MAX_PRJ_TYPE;

	////////////////////////////////////////////
	// m_pActType->asIntArray[][] : activity_struct 시트의 cells(3,2) ~ cells(7,14) 의 값이 들어 있음.
	
	////////////////////////////////////////////
	// 프로젝트 타입관련 정보	
	for (i = 0; i < maxLoop; ++i) { // 프로젝트 타입을 결정한다
		UB += m_ActType.asIntArray[i][0];	// 엑셀 2열의 "발생 확률"

		if (Lb <= probability && probability < UB) {
			pProject->projectType = i;
			break;
		}
		Lb = UB;
	}
	
	Lb = m_ActType.asIntArray[pProject->projectType][2];	// 엑셀 4열의 "최소기간"
	UB = m_ActType.asIntArray[pProject->projectType][3];	// 엑셀 5열의 "최대기간"

	totalDuration = RandomBetween(Lb, UB);
	pProject->duration = totalDuration;
	pProject->endDate = pProject->startAvail + totalDuration - 1;// song??

	Lb = 0;
	UB = 0;
	maxLoop = m_ActType.asIntArray[pProject->projectType][4];//패턴수

	// 패턴 타입 결정
	for (i = 0; i < maxLoop; ++i) {
		UB += m_ActType.asIntArray[pProject->projectType][6 + ((i) * 2)];//1번패턴 확률부터

		if (Lb <= probability && probability < UB) {
			pProject->activityPattern = m_ActType.asIntArray[pProject->projectType][5 + ((i) * 2)];//1번패턴 패턴번호부터
			break;
		}
		Lb = UB;
	}

	//////////////////////////////////////////////////////////////////
	//프로젝트 패턴 관련 정보
	Lb = 0;
	UB = 0;
	maxLoop = m_ActPattern.asIntArray[pProject->activityPattern][0];//활동수 !!! -1 에 주의
	pProject->numActivities = maxLoop;

	// 활동 생성
	for (i = 0; i < maxLoop; ++i) {
		Lb += m_ActPattern.asIntArray[pProject->activityPattern][1 + ((i) * 5)];// !!! -1 에 주의
		UB += m_ActPattern.asIntArray[pProject->activityPattern][2 + ((i) * 5)];// !!! -1 에 주의
		probability = RandomBetween(Lb, UB);
		tempDuration = totalDuration * probability / 100;

		if (tempDuration == 0) {
			tempDuration = 1;
		}

		if (i == 0) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail;
			pProject->activities[i].endDate = pProject->startAvail - 1 + tempDuration;
		}
		else if (i == 1) {
			pProject->activities[i].duration = totalDuration - pProject->activities[0].duration;
			pProject->activities[i].startDate = pProject->activities[0].endDate + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else if (i == 2) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail - 1 + totalDuration - tempDuration + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->activities[2].startDate - tempDuration;
			pProject->activities[i].endDate = pProject->activities[2].startDate - 1;
		}
	}
	return TRUE;
}

// 활동별 투입 인력 생성 및 프로젝트 전체 기대 수익 계산 함수
int CCreator::CalculateHRAndProfit(PROJECT* pProject) {
	int high = 0, mid = 0, low = 0;

	for (int i = 0; i < pProject->numActivities; ++i) {
		int j = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
		if (0 < j && j <= RND_HR_H) {
			pProject->activities[i].highSkill	= 1;
			pProject->activities[i].midSkill	= 0;
			pProject->activities[i].lowSkill	= 0;
		}
		else if (RND_HR_H < j && j <= RND_HR_M) {
			pProject->activities[i].highSkill	= 0;
			pProject->activities[i].midSkill	= 1;
			pProject->activities[i].lowSkill	= 0;
		}
		else {
			pProject->activities[i].highSkill = 0;
			pProject->activities[i].midSkill = 0;
			pProject->activities[i].lowSkill	= 1;
		}
	}

	// 내부 프로젝트는 인력 배정이 틀려진다.
	if (1 == pProject->category)
	{
		pProject->activities[0].highSkill	= pProject->activities[0].highSkill + 1;
		
		pProject->activities[1].highSkill	= pProject->activities[1].highSkill + 1;
		pProject->activities[1].midSkill	= pProject->activities[1].midSkill+1;
		pProject->activities[1].lowSkill	= 1;

	}

	for (int i = 0; i < pProject->numActivities; ++i) {
		high += pProject->activities[i].highSkill* pProject->activities[i].duration;
		mid +=  pProject->activities[i].midSkill * pProject->activities[i].duration;
		low +=  pProject->activities[i].lowSkill * pProject->activities[i].duration;
	}

	return CalculateTotalLaborCost(high, mid, low);
}

// 등급별 투입 인력 계산 및 프로젝트의 수익 생성 함수
double CCreator::CalculateTotalLaborCost(int highCount, int midCount, int lowCount) {
	double highLaborCost	= CalculateLaborCost("H") * highCount;
	double midLaborCost	= CalculateLaborCost("M") * midCount;
	double lowLaborCost	= CalculateLaborCost("L") * lowCount;
	
	double totalLaborCost = highLaborCost + midLaborCost + lowLaborCost;
	return totalLaborCost;
}

// 등급별 투입 인력에 따른 수익 계산 함수
// 수정시 다음도 수정 필요 BOOL CCompany::Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad)
double CCreator::CalculateLaborCost(const std::string& grade) {
	double directLaborCost	= 0;
	double overheadCost	= 0;
	double technicalFee	= 0;
	double totalLaborCost	= 0;

	// 입력된 grade를 대문자로 변환
	char upperGrade = std::toupper(static_cast<unsigned char>(grade[0]));

	switch (upperGrade) {
	case 'H':
		directLaborCost = 50;
		break;
	case 'M':
		directLaborCost = 39;
		break;
	case 'L':
		directLaborCost = 25;
		break;
	default:
		AfxMessageBox(_T("잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."), MB_OK | MB_ICONERROR);
		return 0; // 잘못된 입력 시 함수 종료
	}

	// 간접 : 직접 : 기술 = 6:2:2 = 10 ==> 소숫점이 나오지 않게 10배 키워서 계산한다.
	overheadCost = directLaborCost * 0.6; // 간접 비용 계산
	technicalFee = (directLaborCost + overheadCost) * 0.2; // 기술 비용 계산
	totalLaborCost = directLaborCost + overheadCost + technicalFee; // 총 인건비 계산

	return totalLaborCost;
}


// 대금 지급 조건 생성 함수
void CCreator::CalculatePaymentSchedule(PROJECT* pProject) {

	int paymentType;
	int paymentRatio;
	int totalPayments;

	pProject->firstPayMonth = 1;
	
	// 6주 이하의 짧은 프로젝트는 선금, 잔금만 있다.
	if (pProject->duration <= 6) {
		paymentType = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

		switch (paymentType) {
		case 1:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
			pProject->cashFlows[0] = 30;
			pProject->cashFlows[1] = 70;
			break;
		case 2:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
			pProject->cashFlows[0] = 40;
			pProject->cashFlows[1] = 60;
			break;
		case 3:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
			pProject->cashFlows[0] = 50;
			pProject->cashFlows[1] = 50;
			break;
		}

		pProject->secondPay = pProject->profit - pProject->firstPay;
		totalPayments = 2;
		pProject->secondPayMonth = pProject->duration;
	}

	// 7~12주 사이의 프로젝트는 3회에 걸처셔 받는다.
	else if (pProject->duration <= 12) {
		paymentType = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

		if (paymentType <= 3) {
			paymentRatio = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

			switch (paymentRatio) {
			case 1:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 70;
				break;
			case 2:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
				pProject->cashFlows[0] = 40;
				pProject->cashFlows[1] = 60;
				break;
			case 3:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
				pProject->cashFlows[0] = 50;
				pProject->cashFlows[1] = 50;
				break;
			}

			pProject->secondPay = pProject->profit - pProject->firstPay;
			totalPayments = 2;
			pProject->secondPayMonth = pProject->duration;
		}
		else {
			paymentRatio = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

			if (paymentRatio <= 6) {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 30;
				pProject->cashFlows[2] = 40;
			}
			else {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 40;
				pProject->cashFlows[2] = 30;
			}

			pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;
			totalPayments = 3;
			pProject->secondPayMonth = (int)ceil((double)pProject->duration / 2);
			pProject->finalPayMonth = pProject->duration;
		}
	}

	// 1년 이상의 프로젝트는 3회에 걸처서 받는다.
	else {
		pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
		pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
		pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;

		pProject->cashFlows[0] = 30;
		pProject->cashFlows[1] = 40;
		pProject->cashFlows[2] = 30;

		totalPayments = 3;
		pProject->secondPayMonth = (int)ceil((double)pProject->duration / 2);
		pProject->finalPayMonth = pProject->duration;
	}
	pProject->nCashFlows = totalPayments;
}


void CCreator::Save(CString filename)
{

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

	// 정품 키 값이 들어 있다. 공개하는 프로젝트에는 포함되어 있지 않다. 
	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다.
#ifdef INCLUDE_LIBXL_KET
	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif


	Sheet* projectSheet = nullptr;
	Sheet* dashboardSheet = nullptr;

	if (book->load(filename)) {
		// File exists, check for specific sheets


		for (int i = 0; i < book->sheetCount(); ++i) {
			Sheet* sheet = book->getSheet(i);
			if (std::wcscmp(sheet->name(), L"project") == 0) {
				projectSheet = sheet;
				clearSheet(projectSheet);  // Assuming you have a clearSheet function defined
			}
			if (std::wcscmp(sheet->name(), L"dashboard") == 0) {
				dashboardSheet = sheet;
				clearSheet(dashboardSheet);  // Assuming you have a clearSheet function defined
			}
		}

		if (!projectSheet) {
			projectSheet = book->addSheet(L"project");
		}

		if (!dashboardSheet) {
			dashboardSheet = book->addSheet(L"dashboard");
		}
	}
	else {
		// File does not exist, create new file with sheets
		projectSheet = book->addSheet(L"project");  // Add and assign the 'project' sheet
		dashboardSheet = book->addSheet(L"dashboard");  // Add and assign the 'dashboard' sheet
	}


	write_global_env(book, projectSheet,&m_GlobalEnv);
	write_project_header(book, projectSheet);

	for (int i = 0; i < m_totalProjectNum; i++) {
		write_project_body(book, projectSheet, &(m_pProjects[0][i]));  // Assuming write_project_body is defined
	}

	projectSheet->writeStr(1, 2, L"TatalPrjNum");
	projectSheet->writeNum(1, 3, m_totalProjectNum);

	// Save and release
	book->save(filename);
	book->release();
}

void CCreator::PrintProjectInfo() {

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

	// 정품 키 값이 들어 있다. 공개하는 프로젝트 소스코드에는 포함되어 있지 않다. 
	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다. 필요시 구매해서 사용 바람.
	#ifdef INCLUDE_LIBXL_KET
		book->setKey(_LIBXL_NAME, _LIBXL_KEY);
	#endif

	Sheet* projectSheet = nullptr;
	Sheet* dashboardSheet = nullptr;

	if (book->load(L"d:/new.xlsx")) {
		// File exists, check for specific sheets

		for (int i = 0; i < book->sheetCount(); ++i) {
			Sheet* sheet = book->getSheet(i);
			if (std::wcscmp(sheet->name(), L"project") == 0) {
				projectSheet = sheet;
				clearSheet(projectSheet);  // Assuming you have a clearSheet function defined
			}
			if (std::wcscmp(sheet->name(), L"dashboard") == 0) {
				dashboardSheet = sheet;
				clearSheet(dashboardSheet);  // Assuming you have a clearSheet function defined
			}
		}

		if (!projectSheet) {
			projectSheet = book->addSheet(L"project");
		}

		if (!dashboardSheet) {
			dashboardSheet = book->addSheet(L"dashboard");
		}
	}
	else {
		// File does not exist, create new file with sheets
		projectSheet = book->addSheet(L"project");  // Add and assign the 'project' sheet
		dashboardSheet = book->addSheet(L"dashboard");  // Add and assign the 'dashboard' sheet
	}

	write_project_header(book, projectSheet);

	for (int i = 0; i < m_totalProjectNum; i++) {
		write_project_body(book, projectSheet, &(m_pProjects[0][i]));  // Assuming write_project_body is defined
	}

	// Save and release
	book->save(L"d:/new.xlsx");
	book->release();
}