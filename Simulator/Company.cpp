#include "pch.h"
#include "GlobalEnv.h"
#include "Company.h"
#include "Creator.h"

CCompany::CCompany()
{
	recruitTerm = 8; // (분기 마다 인원 증감 결정 ==> 1000/12주)
}

CCompany::~CCompany()
{
	ClearMemory();
	
	if (m_AllProjects)
		delete[] m_AllProjects;
}

void CCompany::ClearMemory() 
{
	if (m_orderTable[0] != NULL)
	{
		delete[] m_orderTable[0];
		m_orderTable[0] = NULL;
	}

	if (m_orderTable[1] != NULL)
	{
		delete[] m_orderTable[1];
		m_orderTable[1] = NULL;
	}
}

BOOL CCompany::Init(CString fileName)
{
    Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

    // 정품 키 값이 들어 있다. 공개하는 프로젝트에는 포함되어 있지 않다. 
    // 정품 키가 없으면 읽기가 300 컬럼으로 제한된다.
#ifdef INCLUDE_LIBXL_KET
    book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif

    Sheet* projectSheet = nullptr;

    if (book->load(fileName)) {
        // File exists, check for specific sheets
        for (int i = 0; i < book->sheetCount(); ++i) {
            Sheet* sheet = book->getSheet(i);
            if (std::wcscmp(sheet->name(), L"project") == 0) {
                projectSheet = sheet;
            }
        }

        if (!projectSheet) {
            AfxMessageBox(_T("Sheet does not exist"), MB_OK | MB_ICONERROR);
            return FALSE;
        }

    }
    else {
        // File does not exist, create new file with sheets
        AfxMessageBox(_T("File does not exist"), MB_OK | MB_ICONERROR);
        return FALSE;
    }

    read_global_env(book, projectSheet, &m_GlobalEnv);
    m_totalProjectNum = projectSheet->readNum(1, 3);

    m_AllProjects = new PROJECT[m_totalProjectNum];

	// POJECT.runningWeeks 와 _ACTIVITY.activityType;은 파라메터로 전달되지 않는다. 0 으로 초기화 하자	
    memset(m_AllProjects, 0, sizeof(PROJECT) * m_totalProjectNum);

    for (int i = 0; i < m_totalProjectNum; i++)
    {
        read_project_body(book, projectSheet, m_AllProjects + i, i);
    }

    book->release();

    //song MakeOrderTable(fp); 는 필요시TableInit함수 안에 추가하자 
    TableInit();

    return TRUE;
}

void CCompany::TableInit()
{
	//m_orderTable.Resize(2, nWeeks);

	m_doingHR.Resize(3, m_GlobalEnv.maxWeek);
	m_freeHR.Resize(3, m_GlobalEnv.maxWeek);
	m_totalHR.Resize(3, m_GlobalEnv.maxWeek);

	m_doingTable.Resize(11, m_GlobalEnv.maxWeek);
	m_doneTable.Resize(11, m_GlobalEnv.maxWeek);
	m_defferTable.Resize(11, m_GlobalEnv.maxWeek);

	m_incomeTable.Resize(1, m_GlobalEnv.maxWeek);
	m_expensesTable.Resize(1, m_GlobalEnv.maxWeek);

	m_doingHR.Clear();
	m_totalHR.Clear();
	m_freeHR.Clear();	
	
	m_doingTable.Clear();
	m_doneTable.Clear();
	m_defferTable.Clear();

	m_incomeTable.Clear();
	m_expensesTable.Clear();

	m_totalHR[HR_HIG][0] = m_freeHR[HR_HIG][0] = m_GlobalEnv.Hr_Init_H;
	m_totalHR[HR_MID][0] = m_freeHR[HR_MID][0] = m_GlobalEnv.Hr_Init_M;
	m_totalHR[HR_LOW][0] = m_freeHR[HR_LOW][0] = m_GlobalEnv.Hr_Init_L;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_GlobalEnv.ExpenseRate;
	int expenses = (m_GlobalEnv.Hr_Init_H * 50 * rate) + (m_GlobalEnv.Hr_Init_M * 39 * rate) + (m_GlobalEnv.Hr_Init_L * 25 * rate);

	for (int i = 0; i < m_GlobalEnv.maxWeek; i++)
	{
		m_totalHR[HR_HIG][i] = m_GlobalEnv.Hr_Init_H;
		m_totalHR[HR_MID][i] = m_GlobalEnv.Hr_Init_M;
		m_totalHR[HR_LOW][i] = m_GlobalEnv.Hr_Init_L;
		m_expensesTable[0][i] = expenses;
	}
}

// Order table 복구
void CCompany::ReadOrder(FILE* fp)
{
	int orderTableSize = sizeof(int) * m_GlobalEnv.maxWeek * 2;  // 바이트 단위로 크기 계산
	int* temp = new int[m_GlobalEnv.maxWeek * 2];
	ReadDataWithHeader(fp, temp, orderTableSize, TYPE_ORDER);

	// 기존의 m_orderTable이 있었다면 해제
	if (m_orderTable[0] != NULL)
	{
		delete[] m_orderTable[0];
		m_orderTable[0] = NULL;
	}

	if (m_orderTable[1] != NULL)
	{
		delete[] m_orderTable[1];
		m_orderTable[1] = NULL;
	}

	int* order0 = new int[m_GlobalEnv.maxWeek];
	int* order1 = new int[m_GlobalEnv.maxWeek];

	memcpy(order0, (int*)temp, m_GlobalEnv.maxWeek*sizeof(int));
	memcpy(order1, (int*)temp + m_GlobalEnv.maxWeek, m_GlobalEnv.maxWeek * sizeof(int));

	m_orderTable[0] = order0;
	m_orderTable[1] = order1;

	delete temp;
	// 필요 없음 delete order0;
	// 필요 없음 delete order1;

}


// Order table 복구
void CCompany::ReadProject(FILE* fp)
{
	SAVE_TL tl;
	if (fread(&tl, 1, sizeof(tl), fp) != sizeof(tl)) {
		perror("Failed to read header");		
	}

	if (tl.type == TYPE_PROJECT) 
	{
		m_totalProjectNum = tl.length;
		m_AllProjects = new PROJECT[m_totalProjectNum];
	}

	fread(m_AllProjects, sizeof(PROJECT), m_totalProjectNum,  fp);
}


// 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
BOOL CCompany::Decision(int thisWeek ) {

	m_lastDecisionWeek = thisWeek;

	// 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
	if (FALSE == CheckLastWeek(thisWeek))
	{
		//파산		
		return FALSE;
	}

	// 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
	SelectCandidates(thisWeek);

	// 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	// 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
	SelectNewProject(thisWeek);

	//PrintDBData();
	return TRUE;
}



// 완료프로젝트 검사 및 진행프로젝트 업데이트
// 1. 지난 기간의 정보를 이번기간에 복사하고
// 2. 지난 기간에 진행중인 프로젝트중 완료된 것이 있는가?
// 3. 완료된 프로젝트들만 이번기간에서 삭제
BOOL CCompany::CheckLastWeek(int thisWeek)
{
	if (0 == thisWeek) // 첫주는 체크할 지난주가 없음
		return TRUE;

	int nLastProjects = m_doingTable[ORDER_SUM][thisWeek - 1];//지난주에 진행 중이던 프로젝트의 갯수

	for (int i = 0; i < nLastProjects; i++)
	{
		int prjId = m_doingTable[i + 1][thisWeek - 1];
		if (prjId == 0)
			return TRUE;

		PROJECT* project = m_AllProjects + (prjId - 1);

		if (project->category == 0 ) {// 외부프로젝트면
			// 1. payment 를 계산한다. 선금은 시작시 받기로 한다. 조건완료후 1주 후 수금			
			// 2. 지출을 계산한다.
			// 3. 진행중인 프로젝트를 이관해서 기록한다.
			int sum = m_doingTable[ORDER_SUM][thisWeek];
			if (thisWeek < (project->runningWeeks + project->duration - 1)) // ' 아직 안끝났으면
			{
				m_doingTable[sum + 1][thisWeek] = project->ID;// 테이블 크기는 자동으로 변경된다.
				m_doingTable[ORDER_SUM][thisWeek] = sum + 1;
			}
		}
		else // 내부프로젝트
		{	//song !!!
			// 1. 지난주에 종료되었으면 앞으로 받을 금액표 업데이트
			if (project->endDate == (thisWeek - 1))
			{
				int win = ZeroOrOneByProb(project->winProb); // 성공 확율에 따라서 금액을 결정한다.

				if (win) {
					for (int future = thisWeek; future < m_GlobalEnv.maxWeek ; future++) {
						m_incomeTable[0][future] += project->profit;
					}
				}
			}
			else {

				// 2. 진행중이면 내부프로젝트는 이번주부터 시작 가능 으로 표시하고
				// 지금까지 진행된 기간을 runningWeeks 에 한주 증가시켜서 표시함.
				project->startAvail = thisWeek;				
				project->runningWeeks += 1;

				// 내부 프로젝트의 인력 테이블 조정
				RemoveInterProject(project, thisWeek);
			}
		}
	}

	// 자금 현황을 체크하자.
	// 나중에 후회 해도 일단은 편하게 코딩.	

	// 현재 보유중인 현금
	int Cash = m_GlobalEnv.Cash_Init;
	for (int i = 0; i < thisWeek; i++)
	{
		Cash += (m_incomeTable[0][i] - m_expensesTable[0][i]);
	}

	// 이번주 현금은 이상이 없는가?
	if (Cash < 0)// 이번주에 파산
	{
		return FALSE;
	}


	/// 인원 충원을 결정하자.
	
	// 지금부터 채용한계선까지의 수지 차이
	// 이렇게 하면 너무 채용이 많아짐. 
	int temp = m_GlobalEnv.Cash_Init; // 기간까지 필요한 현금 = 필요지출 - 예상수입
	for (int i = 0; i < m_GlobalEnv.recruit; i++)
	{
		temp += m_expensesTable[0][i + thisWeek] - (m_incomeTable[0][i + thisWeek]) ;
	}

	// 보유 현금으로 인원 충원 한계선 이상 유지가 가능하면 충원		
	if (temp < Cash)
	{
		//분기에 한번 꼴로 충원하자.
		if (0 < recruitTerm){ //0 이면 인원 증감 없음
			int win = ZeroOrOneByProb(recruitTerm); // 분기에 한번 충원
			if (win) {
				int i = rand() % 3; /// 고급,중급,초급중 아무나
				AddHR(i, thisWeek + m_GlobalEnv.Hr_LeadTime);// 인원 충원 리드 타임
			}
		}
		
	}

	else 
	{
		temp = 0;// m_GlobalEnv.Cash_Init;
		for (int i = 0; i < m_GlobalEnv.layoff; i++)
		{
			temp += (m_expensesTable[0][i + thisWeek] - m_incomeTable[0][i + thisWeek] );
		}

		if (temp > Cash)
		{
			if (0 < recruitTerm) { //0 이면 인원 증감 없음
				int win = ZeroOrOneByProb(recruitTerm); // 분기에 한번 감원
				if (win) {
					int i = rand() % 3;  //song 인원 감원은 프로젝트 할당 상황을 보고 결정하게 수정해야함.
					RemoveHR(i, thisWeek + m_GlobalEnv.Hr_LeadTime);// 인원 감원 리드 타임
				}
			}
		}
	}
	
	return TRUE;
}

// 이번주에 진행했던!!! 내부프로젝트의 이번주까지의 진행상황을 체크해서 
// 다음주부터 배정되어 있던 인원을 삭제한다.
// 이번주에 진행중이 아니었으면 배정된 인원이 없고, 삭제할 필요도 없다.
void CCompany::RemoveInterProject(PROJECT* project, int thisWeek)
{	
	//project->runningWeeks => 현재가지 진행된 week 수. 0 이면 시작안한 상태
	int runningWeeks = project->runningWeeks;
	int startAvail = project->startAvail;

	//song  !!!! int gap = thisWeek - startAvail  - runningWeeks; 값들간에 문제는 없는지 체크
	// 밀린 시간 = 경과 시간 - 진행한 시간
	// 경과 시간 = 현재 시각 - 시작 가능 시작
	int gap = thisWeek - startAvail - runningWeeks; 

	// 예외를 위해서 메세지 박스를 띄운다. 디버깅용

	// Activity들의 원본은 그대로 두고 복사해온 곳에 밀린만큼 모든 기간을 수정한다.
	int numAct = project->numActivities;

	PACTIVITY pActivity = new ACTIVITY[numAct];
	for (int i = 0; i < numAct; i++)
	{
		pActivity[i] = *project->activities;
		pActivity[i].startDate += gap;
		pActivity[i].endDate += gap;

		int startDate = pActivity[i].startDate;
		int endDate = pActivity[i].endDate;

		for (int j = startDate; j <= endDate; j++)
		{	
			if(j > thisWeek)
			{
				m_doingHR[HR_HIG][j] -= pActivity[i].highSkill;
				m_doingHR[HR_MID][j] -= pActivity[i].midSkill;
				m_doingHR[HR_LOW][j] -= pActivity[i].lowSkill;

				m_freeHR[HR_HIG][j] = m_totalHR[HR_HIG][j] - m_doingHR[HR_HIG][j];
				m_freeHR[HR_MID][j] = m_totalHR[HR_MID][j] - m_doingHR[HR_MID][j];
				m_freeHR[HR_LOW][j] = m_totalHR[HR_LOW][j] - m_doingHR[HR_LOW][j];
			}			
		}
	}

	delete[] pActivity;
}



void CCompany::AddHR(int grade ,int addWeek)
{
	// 충원 / 감원 인원 추가
	// 나머지 기간 업데이트
	// 나머지 기간의 비용 업데이트
	m_totalHR[grade][addWeek] = m_totalHR[grade][addWeek] + 1;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_GlobalEnv.ExpenseRate;
	int expenses = (m_totalHR[0][addWeek] * 50 * rate) + (m_totalHR[1][addWeek] * 39 * rate) + (m_totalHR[2][addWeek] * 25 * rate);

	for (int i = addWeek; i < m_GlobalEnv.maxWeek; i++)
	{
		m_totalHR[HR_HIG][i] = m_totalHR[HR_HIG][addWeek];
		m_totalHR[HR_MID][i] = m_totalHR[HR_MID][addWeek];
		m_totalHR[HR_LOW][i] = m_totalHR[HR_LOW][addWeek];
		m_expensesTable[0][i] = expenses;
	}
}
//
void CCompany::RemoveHR(int grade, int removeWeek)
{
	// 감원 인원 
	// 나머지 기간 업데이트
	// 나머지 기간의 비용 업데이트
	m_totalHR[grade][removeWeek] = m_totalHR[grade][removeWeek] + 1;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_GlobalEnv.ExpenseRate;
	int expenses = (m_totalHR[0][removeWeek] * 50 * rate) + (m_totalHR[1][removeWeek] * 39 * rate) + (m_totalHR[2][removeWeek] * 25 * rate);

	for (int i = removeWeek; i < m_GlobalEnv.maxWeek; i++)
	{
		m_totalHR[HR_HIG][i] = m_totalHR[HR_HIG][removeWeek];
		m_totalHR[HR_MID][i] = m_totalHR[HR_MID][removeWeek];
		m_totalHR[HR_LOW][i] = m_totalHR[HR_LOW][removeWeek];
		m_expensesTable[0][i] = expenses;
	}
}

// song 프로젝트 테이블을 모두 돌면서 order이 이번주인것에서 비교하게 변경하자.

void CCompany::SelectCandidates(int thisWeek)
{	
	for (int i = 0; i< MAX_CANDIDATES; i++)
		m_candidateTable[i] = 0;

	int j = 0;

	for (int i = 0; i < m_totalProjectNum; i++)
	{
		PROJECT* project = m_AllProjects + i;
		if((project->category == 1)) { //내부프로젝트
			if (project->duration > project->runningWeeks) { // 완료되지 않은 내부프로젝트
				if (IsInternalEnoughHR(thisWeek, project)) // 내부 프로젝트 인원 체크			
					m_candidateTable[j++] = project->ID;
			}
		}
		else if(project->orderDate == thisWeek) // 이번주 발생 프로젝트
		{
			if (IsEnoughHR(thisWeek, project)) // 인원 체크			
				m_candidateTable[j++] = project->ID;
		}
	}
}

//
//void CCompany::SelectCandidatesOld(int thisWeek)
//{
//	int lastID = m_orderTable[ORDER_SUM][thisWeek] ;	// 지난달까지 누계
//	int endID = m_orderTable[ORDER_ORD][thisWeek] + lastID;  // 지난달까지 누계 + 이번주 발생 갯수 - 1
//	for(int i=0; i< MAX_CANDIDATES; i++)
//		m_candidateTable[i] = 0;
//
//	int j = 0; 
//	for (int i = lastID; i < endID; i++)
//	{
//		PROJECT* project = m_AllProjects + i;
//
//		if (IsEnoughHR(thisWeek, project)) // 인원 체크
//		{
//			m_candidateTable[j++] = project->ID;
//		}
//	}
//
//
//	// 내부프로젝트에서 후보군을 찾는다.
//	for (int i = 0; i < 1; i++)
//	{
//		PROJECT* project = m_InterProjects + i;
//
//		if (IsEnoughHR(thisWeek, project)) // 인원 체크
//		{
//			m_candidateTable[j++] = project->ID;
//		}
//	}
//}


// 다음주에 내부 프로젝트를 위해 투입할 인원 체크
BOOL CCompany::IsInternalEnoughHR(int thisWeek, PROJECT* project)
{
	//project->runningWeeks => 현재가지 진행된 week 수. 0 이면 시작안한 상태
	int runningWeeks = project->runningWeeks;
	int startAvail = project->startAvail;

	//song  !!!! int gap = thisWeek - startAvail  - runningWeeks; 값들간에 문제는 없는지 체크
	// 밀린 시간 = 경과 시간 - 진행한 시간
	// 경과 시간 = 현재 시각 - 시작 가능 시작
	int gap = thisWeek - startAvail - runningWeeks;

	// 예외를 위해서 메세지 박스를 띄운다. 디버깅용

	// Activity들의 원본은 그대로 두고 복사해온 곳에 밀린만큼 모든 기간을 수정한다.
	int numAct = project->numActivities;
	Dynamic2DArray doingHR = m_doingHR;

	PACTIVITY pActivity = new ACTIVITY[numAct];

	for (int i = 0; i < numAct; i++)
	{
		pActivity[i] = *project->activities;
		pActivity[i].startDate += gap;
		pActivity[i].endDate += gap;

		int startDate = pActivity[i].startDate;
		int endDate = pActivity[i].endDate;

		for (int j = startDate; j <= endDate; j++)
		{
			if (j == thisWeek+1) // 다음주 인원만 고려하자
			{
				doingHR[HR_HIG][j] -= pActivity[i].highSkill;
				doingHR[HR_MID][j] -= pActivity[i].midSkill;
				doingHR[HR_LOW][j] -= pActivity[i].lowSkill;
			}
		}
	}

	delete[] pActivity;

	if (m_totalHR[HR_HIG][thisWeek + 1] < doingHR[HR_HIG][thisWeek + 1])
		return FALSE;
	if (m_totalHR[HR_MID][thisWeek + 1] < doingHR[HR_MID][thisWeek + 1])
		return FALSE;
	if (m_totalHR[HR_LOW][thisWeek + 1] < doingHR[HR_LOW][thisWeek + 1])
		return FALSE;

	return TRUE;
}


BOOL CCompany::IsEnoughHR(int thisWeek, PROJECT* project)
{
	// 원본 인력 테이블을 복사해서 프로젝트 인력을 추가 할 수 있는지 확인한다.
	Dynamic2DArray doingHR = m_doingHR;
		
	// 2중 루프 activity->기간-> 등급업데이트 순서로 activity들을 순서대로 가져온다.
	int numAct = project->numActivities;
	for (int i = 0 ; i < numAct ;i++)
	{
		PACTIVITY pActivity = &(project->activities[i]);
		for (int j = 0; j < pActivity->duration; j++)
		{
			doingHR[HR_HIG][j + pActivity->startDate] += pActivity->highSkill;
			doingHR[HR_MID][j + pActivity->startDate] += pActivity->midSkill;
			doingHR[HR_LOW][j + pActivity->startDate] += pActivity->lowSkill;
		}		
	}

	for (int i = thisWeek; i < m_GlobalEnv.maxWeek; i++) 
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i])
			return FALSE;
			
		if (m_totalHR[HR_MID][i] < doingHR[HR_MID][i])
			return FALSE;

		if (m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
			return FALSE;
	}

	return TRUE;
}

// 후보군들을 선택 정책에 따라서 순서를 변경한다.

// 2차원 배열을 오름차순으로 정렬하는 함수
void sortArrayAscending(int* indexArray, int* valueArray, int size) {
	// 두 배열을 정렬하기 위해 값과 인덱스를 페어로 묶어야 합니다.
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] > valueArray[j]) {
				// 값(value)을 기준으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}

// 2차원 배열을 내림차순으로 정렬하는 함수
void sortArrayDescending(int* indexArray, int* valueArray, int size) {
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] < valueArray[j]) {
				// 값(value)을 기준으로 내림차순으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}

void CCompany::SelectNewProject(int thisWeek)
{	
	int valueArray[MAX_CANDIDATES] = {0, };  // 값 배열
	int j = 0;

	while (m_candidateTable[j] != 0) {

		PROJECT* project;
		int id = m_candidateTable[j];

		project = m_AllProjects + (id - 1);
		if(project->category == 0){// 외부 프로젝트
			valueArray[j] = project->profit;
		}
		else {  //내부 프로젝트
			valueArray[j] = project->profit * 4 * 12 *3;
		}
		j = j + 1;
	}
	
	// 설정된 우선순위대로 프로젝트를 재 배치 한다.
	switch (m_GlobalEnv.selectOrder)
	{
	case 1: // 발생 순서대로
		break;
	case 2:
		sortArrayAscending(m_candidateTable, valueArray, j);	// 금액 내림차순 정렬	
		break;
	case 3:
		sortArrayDescending(m_candidateTable, valueArray, j); // 금액 오름차순 정렬	
		break;
	default : 
		break;
	} 

	int i = 0;
	while (m_candidateTable[i] != 0) {

		if (i > MAX_CANDIDATES) break;

		int id = m_candidateTable[i++];

		PROJECT* project = m_AllProjects + id;

		if (project->startAvail < m_GlobalEnv.maxWeek)
		{
			if (IsEnoughHR(thisWeek, project))
			{
				AddProjectEntry(project, thisWeek);
			}
		}
	}
}

// 모든 체크가 끝나고 호출된다. 
// 단지 변수들만 셑팅하자.
void CCompany::AddProjectEntry(PROJECT* project,  int addWeek)
{	
	//song !!!runningWeeks 함수내 모두 확인 바람
	project->runningWeeks = project->startAvail;

	// HR 정보 업데이트
	// 2중 루프 activity->기간-> 등급업데이트 순서로 activity들을 순서대로 가져온다.
	int numAct = project->numActivities;
	for (int i = 0; i < numAct; i++)
	{
		PACTIVITY pActivity = &(project->activities[i]);
		for (int j = 0; j < pActivity->duration; j++)
		{
			int col = j + pActivity->startDate;
			m_doingHR[HR_HIG][col] += pActivity->highSkill;
			m_doingHR[HR_MID][col] += pActivity->midSkill;
			m_doingHR[HR_LOW][col] += pActivity->lowSkill;

			m_freeHR[HR_HIG][col] = m_totalHR[HR_HIG][col] - m_doingHR[HR_HIG][col];
			m_freeHR[HR_MID][col] = m_totalHR[HR_MID][col] - m_doingHR[HR_MID][col];
			m_freeHR[HR_LOW][col] = m_totalHR[HR_LOW][col] - m_doingHR[HR_LOW][col];
		}
	}

	// 현황판 업데이트
	int sum = m_doingTable[0][addWeek];
	m_doingTable[sum + 1][addWeek] = project->ID;
	m_doingTable[0][addWeek] = sum + 1;

	// 수입 테이블 업데이트. 지출은 인원 관리쪽에서 한다.	
	int incomeDate;

	if (project->runningWeeks <addWeek)
	{
		MessageBox(NULL, _T("m_runningWeeks miss"), _T("Error"), MB_OK | MB_ICONERROR);
	}
	incomeDate = project->runningWeeks + project->firstPayMonth;	// 선금 지급일
	m_incomeTable[0][incomeDate] += project->firstPay;
	
	incomeDate = project->runningWeeks + project->secondPayMonth;	// 2차 지급일
	m_incomeTable[0][incomeDate] += project->secondPay;

	incomeDate = project->runningWeeks + project->finalPayMonth;	// 3차 지급일
	m_incomeTable[0][incomeDate] += project->finalPay;
}


// dash boar 용 배열들의 크기 조절	
//void CCompany::AllTableInit(int nWeeks)
//{
//	m_orderTable = Newallocate2DArray(2, nWeeks);
//
//	m_doingHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);
//	m_freeHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);
//	m_totalHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);
//
//	m_doingTable.Resize(11, nWeeks);
//	m_doneTable.Resize(11, nWeeks);
//	m_defferTable.Resize(11, nWeeks);
//
//	m_incomeTable = Newallocate2DArray(1, nWeeks);
//	m_expensesTable = Newallocate2DArray(1, nWeeks);
//
//
//	// 이건 충원이나 감원쪽에서 필요시 다시 수정하게 된다.	
//	m_totalHR[HR_HIG][0] = m_freeHR[HR_HIG][0] = m_pGlobalEnv->Hr_Init_H;
//	m_totalHR[HR_MID][0] = m_freeHR[HR_MID][0] = m_pGlobalEnv->Hr_Init_M;
//	m_totalHR[HR_LOW][0] = m_freeHR[HR_LOW][0] = m_pGlobalEnv->Hr_Init_L;
//
//	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
//	double rate = m_pGlobalEnv->ExpenseRate;
//	int expenses = (m_pGlobalEnv->Hr_Init_H * 50* rate) + (m_pGlobalEnv->Hr_Init_M * 39* rate) + (m_pGlobalEnv->Hr_Init_L * 25 * rate);
//
//	for (int i = 0; i < nWeeks + ADD_HR_SIZE; i++)
//	{
//		m_totalHR[HR_HIG][i] = m_pGlobalEnv->Hr_Init_H;
//		m_totalHR[HR_MID][i] = m_pGlobalEnv->Hr_Init_M;
//		m_totalHR[HR_LOW][i] = m_pGlobalEnv->Hr_Init_L;
//		m_expensesTable[0][i] = expenses;
//	}
//}

int CCompany::CalculateFinalResult() 
{
	int result = m_GlobalEnv.Cash_Init;

	for (int i = 0; i < m_lastDecisionWeek; i++)
	{
		result += (m_incomeTable[0][i]- m_expensesTable[0][i]);
	}
	
/*	필요시 다음과 같이 처리하자.
	int tempTotalIncome = m_pGlobalEnv->Cash_Init;
	int tempOutcome = m_expensesTable[0][10];
	int tempTetoalOutcome = (tempOutcome*144);
	
	int cols = m_debugInfo.getCols();
	for (int i = 0; i < cols; i++)
	{		
		tempTotalIncome += m_debugInfo[1][i];
	}
	int tempResult = tempTotalIncome- tempTetoalOutcome;
	*/
	//return result;
	//return tempResult; // 기대수익?? 포함 (수주 수익 포함)	
	return result;
}

// 기간내 총 매출
int CCompany::CalculateTotalInCome()
{
	int result = 0;

	for (int i = 0; i < m_lastDecisionWeek; i++)
	{
		result += m_incomeTable[0][i];
	}

	/*	필요시 다음과 같이 처리하자.
	int tempTotalIncome = m_pGlobalEnv->Cash_Init;
	int tempOutcome = m_expensesTable[0][10];
	int tempTetoalOutcome = (tempOutcome*144);

	int cols = m_debugInfo.getCols();
	for (int i = 0; i < cols; i++)
	{
	tempTotalIncome += m_debugInfo[1][i];
	}
	int tempResult = tempTotalIncome- tempTetoalOutcome;
	*/
	//return result;
	//return tempResult; // 기대수익?? 포함 (수주 수익 포함)	
	return result;
}

void CCompany::PrintCompanyResualt()
{
	m_XlFileName = _T("d:\\test\\newresualt.xlsx");
	CString strSheetName = _T("result");

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	
	book->setKey(L"Jiho Song", L"windows-252327000bc9ee046eb86665a2ybo1s890f4fde5");

	Sheet* resultSheet = nullptr;

	if (book->load((LPCWSTR)m_XlFileName)) {

		for (int i = 0; i < book->sheetCount(); ++i) {
			Sheet* sheet = book->getSheet(i);
			if (std::wcscmp(sheet->name(), strSheetName) == 0) {
				resultSheet = sheet;
				clearSheet(resultSheet);  // Assuming you have a clearSheet function defined
			}
		}

		if (!resultSheet) {
			resultSheet = book->addSheet(strSheetName);
		}
	}
	else {
		// File does not exist, create new file with sheets
		resultSheet = book->addSheet(strSheetName);
	}

	write_CompanyInfo(book, resultSheet);

	// Save and release
	book->save((LPCWSTR)m_XlFileName);
	book->release();
}

void CCompany::write_CompanyInfo(Book* book, Sheet* sheet)
{
	int index = 1;// 엑셀의 2행 부터 기록 => 행의 시작이 0번 인덱스

	sheet->writeStr(index++, 0, _T("주"));
	sheet->writeStr(index++, 0, _T("누계"));
	sheet->writeStr(index++, 0, _T("발주"));


	for (int col = 0; col < m_GlobalEnv.SimulationWeeks; ++col) {
		sheet->writeNum(1, col+1, col);
		sheet->writeNum(2, col+1, m_orderTable[0][col]);
		sheet->writeNum(3, col+1, m_orderTable[1][col]);
	}


	index = 5; // 엑셀의 6행 부터 기록

	draw_outer_border(book, sheet, index+1, 0, index+3 , m_GlobalEnv.SimulationWeeks, BORDERSTYLE_THIN, COLOR_BLACK);

	sheet->writeStr(index++, 0, _T("투입"));
	sheet->writeStr(index++, 0, _T("H"));
	sheet->writeStr(index++, 0, _T("M"));
	sheet->writeStr(index++, 0, _T("L"));

	index = 10; // 엑셀의 11행 부터 기록
	sheet->writeStr(index++, 0, _T("여유"));
	sheet->writeStr(index++, 0, _T("H"));
	sheet->writeStr(index++, 0, _T("M"));
	sheet->writeStr(index++, 0, _T("L"));

	index = 15; // 엑셀의 16행 부터 기록
	sheet->writeStr(index++, 0, _T("총원"));
	sheet->writeStr(index++, 0, _T("H"));
	sheet->writeStr(index++, 0, _T("M"));
	sheet->writeStr(index++, 0, _T("L"));


	for (int col = 0; col < m_GlobalEnv.SimulationWeeks; ++col) {

		index = 6; // 엑셀의 7행 부터 기록 ("투입");
		sheet->writeNum(index++, col + 1, m_doingHR[HR_HIG][col]);
		sheet->writeNum(index++, col + 1, m_doingHR[HR_MID][col]);
		sheet->writeNum(index++, col + 1, m_doingHR[HR_LOW][col]);

		index = 11; // 엑셀의 12행 부터 기록 ("여유")
		sheet->writeNum(index++, col + 1, m_freeHR[HR_HIG][col]);
		sheet->writeNum(index++, col + 1, m_freeHR[HR_MID][col]);
		sheet->writeNum(index++, col + 1, m_freeHR[HR_LOW][col]);


		index = 16; // 엑셀의 17행 부터 기록 ("총원")
		sheet->writeNum(index++, col + 1, m_totalHR[HR_HIG][col]);
		sheet->writeNum(index++, col + 1, m_totalHR[HR_MID][col]);
		sheet->writeNum(index++, col + 1, m_totalHR[HR_LOW][col]);
	}




	//// First row settings
	//sheet->writeNum(posY, posX++, pProject->category);
	//sheet->writeNum(posY, posX++, pProject->ID);
	//sheet->writeNum(posY, posX++, pProject->duration);
	//sheet->writeNum(posY, posX++, pProject->startAvail);
	//sheet->writeNum(posY, posX++, pProject->endDate);
	//sheet->writeNum(posY, posX++, pProject->orderDate);
	//sheet->writeNum(posY, posX++, pProject->profit);
	//sheet->writeNum(posY, posX++, pProject->experience);
	//sheet->writeNum(posY, posX++, pProject->winProb);



}