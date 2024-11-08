
// SimulatorDlg.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "Simulator.h"
#include "SimulatorDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#include "Creator.h"
#include "Company.h"

#include <xlnt/xlnt.hpp>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CSimulatorDlg 대화 상자


IMPLEMENT_DYNAMIC(CSimulatorDlg, CDialogEx);

CSimulatorDlg::CSimulatorDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SIMULATOR_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = nullptr;
		
	m_pGlobalEnv = new GLOBAL_ENV; //song delete는 소멸자에서
	
}

CSimulatorDlg::~CSimulatorDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 null로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;

	if (m_pGlobalEnv)
		delete m_pGlobalEnv;
}

void CSimulatorDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CSimulatorDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()	
	ON_BN_CLICKED(IDC_CREATE_PROJECT, &CSimulatorDlg::OnBnClickedCreateProject)
	ON_BN_CLICKED(IDC_SIMULATION, &CSimulatorDlg::OnBnClickedSimulation)
	ON_BN_CLICKED(IDC_CLEAR_XL, &CSimulatorDlg::OnBnClickedClearXl)
	ON_BN_CLICKED(IDC_ONLY_RUN, &CSimulatorDlg::OnBnClickedOnlyRun)	
	ON_BN_CLICKED(IDC_CHANGE_FOLDER, &CSimulatorDlg::OnBnClickedChangeFolder)
END_MESSAGE_MAP()


// CSimulatorDlg 메시지 처리기

BOOL CSimulatorDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.

	int SimulationWeeks = 52 * 5;		// 52주 x 5년	
	int hr_h = 3;
	int hr_m = 2;
	int hr_l = 1;

	m_pGlobalEnv->SimulationWeeks = SimulationWeeks;
	SetDlgItemInt(IDC_SIMWEEKS, m_pGlobalEnv->SimulationWeeks);

	m_pGlobalEnv->maxWeek = SimulationWeeks + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	SetDlgItemInt(IDC_MAX_WEEKS, m_pGlobalEnv->maxWeek);

	m_pGlobalEnv->WeeklyProb = 1.25;
	SetDlgItemText(IDC_WEEKLY_PROB, _T("1.25"));

	m_pGlobalEnv->Hr_Init_H = hr_h;
	SetDlgItemInt(IDC_HR_H, m_pGlobalEnv->Hr_Init_H);
	
	m_pGlobalEnv->Hr_Init_M = hr_m;
	SetDlgItemInt(IDC_HR_M, m_pGlobalEnv->Hr_Init_M);

	m_pGlobalEnv->Hr_Init_L = hr_l;
	SetDlgItemInt(IDC_HR_L, m_pGlobalEnv->Hr_Init_L);

	m_pGlobalEnv->Hr_LeadTime = 4;
	SetDlgItemInt(IDC_LEAD_TIME, m_pGlobalEnv->Hr_LeadTime);

	m_pGlobalEnv->Cash_Init = (HI_HR_COST * hr_h + MI_HR_COST * hr_m + LO_HR_COST * hr_l) * 4 * 12 * 1.2; //인원수 대비 12개월;
	SetDlgItemInt(IDC_CASH, m_pGlobalEnv->Cash_Init);

	m_pGlobalEnv->ProblemCnt = 100;
	SetDlgItemInt(IDC_PROBLEM_CNT, m_pGlobalEnv->ProblemCnt);

	m_pGlobalEnv->selectOrder = 1;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로
	SetDlgItemInt(IDC_ORDER, m_pGlobalEnv->selectOrder);

	m_pGlobalEnv->recruit = 24;  // 작을수록 공격적인 인원 충원(해당 주 만큼을 유지 할 수 있는 이익 잉여가 쌓이면 충원) 
	SetDlgItemInt(IDC_RECRUIT, m_pGlobalEnv->recruit);

	m_pGlobalEnv->layoff = 0;  // 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지
	SetDlgItemInt(IDC_LAY_OFF, m_pGlobalEnv->layoff);

	m_pGlobalEnv->ExpenseRate = 1.2;
	SetDlgItemText(IDC_EXP_RATE, _T("1.2"));

	m_pGlobalEnv->recruitTerm = 12; // 
	SetDlgItemInt(IDC_RECUIT_TERM, m_pGlobalEnv->recruitTerm);

	srand(static_cast<unsigned int>(time(nullptr)));
	CString m_strFolderPath = _T("c:/ahnLab/");
	SetDlgItemText(IDC_SAVE_PATH, m_strFolderPath);

	// 현재 시간을 가져와 원하는 형식으로 파일명을 생성합니다
	CTime currentTime = CTime::GetCurrentTime();
	CString fileName;
	fileName.Format(_T("%02d%02d%02d%02d_save.xlsm"),
		currentTime.GetMonth(),
		currentTime.GetDay(),
		currentTime.GetHour(),
		currentTime.GetMinute());

	SetDlgItemText(IDC_SAVE_FILE, fileName);

	SetDlgItemInt(IDC_PROBLEM_CNT,100);
	

	m_Progress = (CProgressCtrl*)GetDlgItem(IDC_PROGRESS);

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CSimulatorDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CSimulatorDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CSimulatorDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

void CSimulatorDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CSimulatorDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CSimulatorDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CSimulatorDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 애플리케이션을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}


void CSimulatorDlg::GetMainParameters(ALL_ACT_TYPE* act, ALL_ACTIVITY_PATTERN* pattern)
{
	//PALL_ACT_TYPE pActType = new ALL_ACT_TYPE;
	int actTemp[] = {	
		50,  50,  2,  4, 2, 0, 60, 1, 40, 0, 0, 0, 0,
		20,  70,  5, 12, 2, 2, 50, 3, 50, 0, 0, 0, 0,
		20,  90, 13, 26, 2, 4, 50, 5, 50, 0, 0, 0, 0,
		 8,  98, 27, 52, 2, 4, 40, 5, 60, 0, 0, 0, 0,
		 2, 100, 53, 80, 2, 4, 30, 5, 70, 0, 0, 0, 0 };
	//memcpy(pActType, actTemp, 5*9);

	//PALL_ACTIVITY_PATTERN pActPattern = new ALL_ACTIVITY_PATTERN;
	int patternTemp[] = { 
		1, 100, 100,  0, 30, 70, 0, 0,  0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0, 0, 0,
		1, 100, 100, 80, 20,  0, 0, 0,  0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0, 0, 0,
		2,  10,  20, 60, 40,  0, 0, 0, 10, 80, 10, 0, 0, 0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0, 0, 0,
		2,  20,  30, 80, 20,  0, 0, 0, 10, 70, 20, 0, 0, 0,  0,  0, 0, 0, 0,  0,  0, 0, 0, 0, 0, 0,
		3,  20,  30, 70, 30,  0, 0, 0, 10, 60, 30, 0, 0, 0, 60, 40, 0, 0, 0,  0,  0, 0, 0, 0, 0, 0,
		4,  20,  30, 80, 20,  0, 0, 0, 10, 60, 30, 0, 0, 0, 50, 50, 0, 0, 0, 40, 60, 0, 0, 0, 0, 0 };
	//memcpy(pActPattern, patternTemp, 6*26);

	*act = *((ALL_ACT_TYPE*)actTemp);
	*pattern = *((ALL_ACTIVITY_PATTERN*)patternTemp);

	m_pGlobalEnv->SimulationWeeks = GetDlgItemInt(IDC_SIMWEEKS);
	m_pGlobalEnv->maxWeek = GetDlgItemInt(IDC_MAX_WEEKS);

	CString strTemp; 
	GetDlgItemText(IDC_WEEKLY_PROB, strTemp);
	m_pGlobalEnv->WeeklyProb = _wtof(strTemp.GetString());

	m_pGlobalEnv->Hr_Init_H = GetDlgItemInt(IDC_HR_H);
	m_pGlobalEnv->Hr_Init_M = GetDlgItemInt(IDC_HR_M);
	m_pGlobalEnv->Hr_Init_L = GetDlgItemInt(IDC_HR_L);
	m_pGlobalEnv->Hr_LeadTime = GetDlgItemInt(IDC_LEAD_TIME);

	m_pGlobalEnv->Cash_Init = GetDlgItemInt(IDC_CASH);
	m_pGlobalEnv->ProblemCnt = GetDlgItemInt(IDC_PROBLEM_CNT);
	m_pGlobalEnv->selectOrder = GetDlgItemInt(IDC_ORDER);

	m_pGlobalEnv->recruit = GetDlgItemInt(IDC_RECRUIT);
	m_pGlobalEnv->layoff = GetDlgItemInt(IDC_LAY_OFF);
	m_pGlobalEnv->recruitTerm = GetDlgItemInt(IDC_RECUIT_TERM);

	GetDlgItemText(IDC_EXP_RATE, strTemp);
	m_pGlobalEnv->ExpenseRate = _wtof(strTemp.GetString());
	
	GetDlgItemText(IDC_COMMENT,m_strComment);

}

void CSimulatorDlg::OnlyRun(CString strFileName, CString strInSheetName, CString strOutSheetName)
{	
	ALL_ACT_TYPE* actTemp = new ALL_ACT_TYPE;
	ALL_ACTIVITY_PATTERN* patternTemp = new ALL_ACTIVITY_PATTERN;
	GetMainParameters(actTemp, patternTemp);

	CCompany* company = new CCompany;
	company->Init(strFileName, strInSheetName);
	company->m_GlobalEnv.layoff = 20;
	company->m_GlobalEnv.recruit = 20;

	int k = 0;
	while (k < m_pGlobalEnv->SimulationWeeks)
	{
		if (FALSE == company->Decision(k))  // j번째 기간에 결정해야 할 일들
			k = 9999;
		k++;
	}
	company->PrintCompanyResualt(strFileName, strOutSheetName);

	delete actTemp;
	delete patternTemp;
	delete company;

	return;
}


void CSimulatorDlg::MakeProjectAndRun(CString strFileName, CString strInSheetName, CString strOutSheetName)
{
	ALL_ACT_TYPE* actTemp = new ALL_ACT_TYPE;
	ALL_ACTIVITY_PATTERN* patternTemp = new ALL_ACTIVITY_PATTERN;
	GetMainParameters(actTemp, patternTemp);

	CCreator Creator;
	Creator.Init(m_pGlobalEnv, actTemp, patternTemp);
	Creator.Save(strFileName, strInSheetName);

	CCompany* company = new CCompany;
	company->Init(strFileName, strInSheetName);
	
	int k = 0;
	while (k < m_pGlobalEnv->SimulationWeeks)
	{
		if (FALSE == company->Decision(k))  // j번째 기간에 결정해야 할 일들
			k = 9999; 
		//company->PrintCompanyResualt();
		k++;
	}
	company->PrintCompanyResualt(strFileName, strOutSheetName);

	delete actTemp;
	delete patternTemp;
	delete company;

	return;
}


void CSimulatorDlg::OnBnClickedCreateProject()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	//CString strFileName = _T("d:/test/imsi001.xlsm");
	MakePath();
	CString strInSheetName;
	CString strOutSheetName;
	CString strResultSheetName = _T("result");

	

	m_Progress->SetRange(0, m_ProblemCnt*2);
	m_Progress->SetPos(0);
	m_Progress->SetStep(1);	

	for (int i  = 0; i < m_ProblemCnt; i++)
	{		
		strInSheetName.Format(_T("In%03d"),i);
		strOutSheetName.Format(_T("Out%03d"), i);

		m_Progress->SetDlgItemTextW(IDC_PROGRESS, L"Make and run");
		MakeProjectAndRun(m_strFileName, strInSheetName, strOutSheetName);
		m_Progress->StepIt();
		m_Progress->Invalidate();
		
		m_Progress->SetDlgItemTextW(IDC_PROGRESS, L"Make Result");
		MakeResult(m_strFileName, strResultSheetName, strOutSheetName, i);
		m_Progress->StepIt();
		m_Progress->Invalidate();
	}	
	m_Progress->SetPos(0);
}
void CSimulatorDlg::MakeResult(CString strFileName, CString  strResultSheetName, CString  strOutSheetName, int num)
{
	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	
	Sheet* resultSheet = nullptr;
	Sheet* outSheet = nullptr;
#ifdef INCLUDE_LIBXL_KET	
	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif

	if (book->load((LPCWSTR)strFileName)) {

		int sheetCount = book->sheetCount();

		for (int i = 0; i < sheetCount; ++i) 
		{
			Sheet* sheet = book->getSheet(i);

			if (sheet) {
				if (wcscmp(sheet->name(), (LPCWSTR)strResultSheetName) == 0) {
					resultSheet = sheet;
				}
				else if (wcscmp(sheet->name(), (LPCWSTR)strOutSheetName) == 0) {
					outSheet = sheet;
				}
			}
			if (resultSheet && outSheet) break; // 둘 다 찾았으면 루프 종료
		}

		// outSheet가 없으면 에러 처리
		if (!outSheet ) {
			AfxMessageBox(_T("Result sheet not found."));
			book->release();
			return;
		}

		// resultSheet가 없으면 새로 생성
		if (!resultSheet) {
			resultSheet = book->addSheet((LPCWSTR)strResultSheetName);
			if (!resultSheet) {
				AfxMessageBox(_T("Failed to create new result sheet."));
				book->release();
				return;
			}
		}

		// 여기에 resultSheet와 outSheet를 사용한 작업 수행
		if (num==0) {
			resultSheet->writeStr(0, 0, _T("순번"));
			resultSheet->writeStr(0, 1, _T("개월"));
			resultSheet->writeStr(0, 2, _T("S금액"));
			resultSheet->writeStr(0, 3, _T("E금액"));
			resultSheet->writeStr(0, 4, _T("차액"));
			resultSheet->writeStr(0, 5, _T("S인원"));
			resultSheet->writeStr(0, 6, _T("E인원"));
			resultSheet->writeStr(0, 7, _T("증감"));
			resultSheet->writeStr(0, 8, m_strComment);
		}

		resultSheet->writeNum(num + 1, 0, num);

		int lastWeek = outSheet->readNum(0, 1);//종료일
		resultSheet->writeNum(num + 1, 1, lastWeek);

		if (155 <= lastWeek)
		{
			lastWeek = 155;
		}
		
		int SMoney = outSheet->readNum(32, 1);// S금액
		int EMoney = outSheet->readNum(34, lastWeek);// E금액
		int Profit = EMoney - SMoney;

		int SHR = outSheet->readNum(16, 1);// S인원
		SHR += outSheet->readNum(17, 1);// S인원
		SHR += outSheet->readNum(18, 1);// S인원

		int EHR = outSheet->readNum(16, lastWeek);// E인원
		EHR += outSheet->readNum(17, lastWeek);// E인원
		EHR += outSheet->readNum(18, lastWeek);// E인원

		int TotaHR = EHR - SHR;// 최종 인원, 인원 증감


		resultSheet->writeNum(num + 1, 2, SMoney);
		resultSheet->writeNum(num + 1, 3, EMoney);
		resultSheet->writeNum(num + 1, 4, Profit);

		resultSheet->writeNum(num + 1, 5, SHR);
		resultSheet->writeNum(num + 1, 6, EHR);
		resultSheet->writeNum(num + 1, 7, TotaHR);
		
		// 변경사항 저장
		if (!book->save((LPCWSTR)strFileName)) {
			AfxMessageBox(_T("Failed to save the file."));
		}
	}
	else {
		// 파일이 존재하지 않거나 로드 실패
		AfxMessageBox(_T("Failed to load the file."));
		return;
	}

	book->release();
}


void CSimulatorDlg::OnBnClickedSimulation()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

	MakePath();

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	

	// 정품 키 값이 들어 있다. 공개하는 프로젝트에는 포함되어 있지 않다. 
	// 정품 키가 없으면 읽기가 300 컬럼으로 제한된다.
#ifdef INCLUDE_LIBXL_KET	
	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif	

	//CString strFileName = _T("d:/test/111.xlsx");
	Sheet* resultSheet = nullptr;

	if (book->load((LPCWSTR)m_strFileName)) {
		resultSheet = book->getSheet(0);			
		clearSheet(resultSheet);  // Assuming you have a clearSheet function defined		
	}
	else {
		// File does not exist, create new file with sheets
		resultSheet = book->addSheet(_T("box"));
	}

	draw_outer_border(book, resultSheet, 1, 1, 10, 20, BORDERSTYLE_THIN, COLOR_BLACK);
	
	// Save and release
	book->save((LPCWSTR)m_strFileName);
	book->release();

}


void CSimulatorDlg::OnBnClickedClearXl()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.	
	MakePath();

	//CString strFileName = _T("d:/test/imsi000.xlsm");
	CString strOutSheetName;

	Book* book = xlCreateXMLBook();  // Use xlCreateBook() for xls format	
#ifdef INCLUDE_LIBXL_KET	
	book->setKey(_LIBXL_NAME, _LIBXL_KEY);
#endif	
	if (book->load((LPCWSTR)m_strFileName))
	{
		for (int i = 0; i < m_ProblemCnt; i++)
		{
			strOutSheetName.Format(_T("Out%03d"), i);
			// 시트 이름과 일치하는 모든 시트 찾아 삭제
			int sheetCount = book->sheetCount();
			for (int j = sheetCount - 1; j >= 0; j--)  // 역순으로 순회
			{
				Sheet* sheet = book->getSheet(j);
				if (sheet && wcscmp(sheet->name(), (LPCWSTR)strOutSheetName) == 0)
				{
					book->delSheet(j);
				}
			}
		}
		// 변경사항 저장
		if (!book->save((LPCWSTR)m_strFileName))
		{
			AfxMessageBox(_T("Failed to save the file."));
		}
	}
	else
	{
		AfxMessageBox(_T("Failed to load the file."));
	}

	// 리소스 해제
	book->release();
}


void CSimulatorDlg::OnBnClickedOnlyRun()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

	MakePath();

	CString strInSheetName;
	CString strOutSheetName;
	CString strResultSheetName = _T("result");;

	for (int i = 0; i < m_ProblemCnt; i++)
	{
		strInSheetName.Format(_T("In%03d"), i);
		strOutSheetName.Format(_T("Out%03d"), i);

		OnlyRun(m_strFileName, strInSheetName, strOutSheetName);
		MakeResult(m_strFileName, strResultSheetName, strOutSheetName, i);
	}
}



void CSimulatorDlg::OnBnClickedChangeFolder()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	BROWSEINFO BrInfo;
	TCHAR szBuffer[512];                                      // 경로저장 버퍼 

	::ZeroMemory(&BrInfo, sizeof(BROWSEINFO));
	::ZeroMemory(szBuffer, 512);

	BrInfo.hwndOwner = GetSafeHwnd();
	BrInfo.lpszTitle = _T("파일이 저장될 폴더를 선택하세요");
	BrInfo.ulFlags = BIF_NEWDIALOGSTYLE | BIF_EDITBOX | BIF_RETURNONLYFSDIRS;
	LPITEMIDLIST pItemIdList = ::SHBrowseForFolder(&BrInfo);

	if (pItemIdList != NULL)
	{
		// 폴더가 선택되었을 때
		if (::SHGetPathFromIDList(pItemIdList, szBuffer))
		{
			CString str(szBuffer); 
			// 백슬래시를 슬래시로 변경
			str.Replace(_T("\\"), _T("/"));
			str.Append(_T("/"));
			SetDlgItemText(IDC_SAVE_PATH, str);
		}

		// ITEMIDLIST 메모리 해제
		CoTaskMemFree(pItemIdList);
	}	
}


void CSimulatorDlg::MakePath()
{
	CString strTemp;
	GetDlgItemText(IDC_SAVE_FILE, strTemp);

	GetDlgItemText(IDC_SAVE_PATH, m_strFileName);
	
	CString strPath = m_strFileName;
	strPath = m_strFileName;
	strPath.Replace(_T("/") , _T("\\"));
	CreateDirectoryRecursively(strPath);

	m_strFileName.Append(strTemp);
	
	m_ProblemCnt = GetDlgItemInt(IDC_PROBLEM_CNT);
}

BOOL CSimulatorDlg::CreateDirectoryRecursively(const CString& strPath)
{
	CString strTemp = strPath;

	// 경로의 끝에 있는 '/' 또는 '\' 제거
	strTemp.TrimRight(_T("\\/"));

	// 경로가 이미 존재하는지 확인
	if (GetFileAttributes(strTemp) != INVALID_FILE_ATTRIBUTES)
	{
		return TRUE; // 이미 존재하면 성공으로 간주
	}

	// 상위 디렉토리 생성 (재귀적으로)
	int nFound = strTemp.ReverseFind(_T('\\'));
	if (nFound != -1)
	{
		if (!CreateDirectoryRecursively(strTemp.Left(nFound)))
		{
			return FALSE; // 상위 디렉토리 생성 실패
		}
	}

	// 현재 디렉토리 생성
	if (!CreateDirectory(strTemp, NULL))
	{
		DWORD dwError = GetLastError();
		if (dwError != ERROR_ALREADY_EXISTS)
		{
			// 이미 존재하는 경우가 아닌 다른 오류라면 실패
			return FALSE;
		}
	}

	return TRUE;
}

