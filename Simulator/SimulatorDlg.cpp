
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
	// 우선은 동일한 랜덤값으로 프로젝트 생성을 검증해 놓고 
	// 이 이후에 사용하자. 
	// srand(static_cast<unsigned int>(time(nullptr)));

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


void CSimulatorDlg::DefaultParameters(ALL_ACT_TYPE* act, ALL_ACTIVITY_PATTERN* pattern)
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

	int SimulationWeeks = 52 * 5;		// 52주 x 5년
	int hr_h = 3;
	int hr_m = 2;
	int hr_l = 1;

	m_pGlobalEnv->SimulationWeeks = SimulationWeeks;
	m_pGlobalEnv->maxWeek = SimulationWeeks + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	m_pGlobalEnv->WeeklyProb = 1.25;
	m_pGlobalEnv->Hr_Init_H = hr_h;
	m_pGlobalEnv->Hr_Init_M = hr_m;
	m_pGlobalEnv->Hr_Init_L = hr_l;
	m_pGlobalEnv->Hr_LeadTime = 4;
	m_pGlobalEnv->Cash_Init = (50 * hr_h + 39 * hr_m + 25 * hr_l) * 4 * 12 * 1.2; //인원수 대비 12개월;
	m_pGlobalEnv->ProblemCnt = 100;
	m_pGlobalEnv->selectOrder = 1;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로
	m_pGlobalEnv->recruit = 4 * 6;		// 충원에 필요한 운영비 (몇주분량인가?)
	m_pGlobalEnv->layoff = 0;			// 감원에 필요한 운영비 (몇주분량인가?)

	m_pGlobalEnv->ExpenseRate = 1.2;
	m_pGlobalEnv->selectOrder = 0; //선택 순서  1: 먼저 발생한 순서대로 2 : 금액이 큰 순서대로 3 : 금액이 작은 순서대로

	m_pGlobalEnv->recruit = 160;  // 작을수록 공격적인 인원 충원 144 : 시뮬레이션 끝까지 충원 없음
	m_pGlobalEnv->layoff = 0;  // 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지
}



void CSimulatorDlg::OnBnClickedCreateProject()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

	ALL_ACT_TYPE* actTemp = new ALL_ACT_TYPE;
	ALL_ACTIVITY_PATTERN* patternTemp = new ALL_ACTIVITY_PATTERN;
	DefaultParameters(actTemp, patternTemp);

	CString strFileName = _T("d:/test/test.xlsx");

	//CCreator Creator; 
	//Creator.Init(m_pGlobalEnv, actTemp, patternTemp);	
	//Creator.Save(strFileName);
	
	CCompany* company = new CCompany;
	company->Init(strFileName);
	//company->ReInit();

	int k = 0;
	while (k < m_pGlobalEnv->SimulationWeeks)
	{
		//company->PrintCompanyResualt();
		if (FALSE == company->Decision(k))  // j번째 기간에 결정해야 할 일들		
			k = 9999; //m_pGlobalEnv->SimulationWeeks + 1;

		k++;
	}
	company->PrintCompanyResualt();

	delete actTemp;
	delete patternTemp;
	delete company;

	return ;
}



void CSimulatorDlg::OnBnClickedSimulation()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

	ALL_ACT_TYPE* actTemp = new ALL_ACT_TYPE;
	ALL_ACTIVITY_PATTERN* patternTemp = new ALL_ACTIVITY_PATTERN;
	DefaultParameters(actTemp, patternTemp);

#define _PRINT_WIDTH	13
#define _PRINT_CNT	1

	int iLoop = _PRINT_CNT;

	int* lastResult = NULL;

	lastResult = new int[iLoop * _PRINT_WIDTH];

	for (int loop = 0; loop < iLoop; loop++)
	{
		CCreator Creator;
		Creator.Init(m_pGlobalEnv, actTemp, patternTemp);

		CString prarmFile;
		prarmFile =_T("d:\\test\\newresualt.ahn");

		Creator.Save(prarmFile);

		CCompany* company = new CCompany;
		company->Init(prarmFile);
		//company->TableInit();

		int k = 0;
		while (k < m_pGlobalEnv->SimulationWeeks)
		{
			if (FALSE == company->Decision(k))  // j번째 기간에 결정해야 할 일들		
				k = 9999; //m_pGlobalEnv->SimulationWeeks + 1;

			k++;
		}
		company->PrintCompanyResualt();
	}

	delete actTemp;
	delete patternTemp;
	delete lastResult;
}
