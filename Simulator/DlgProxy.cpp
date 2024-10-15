
// DlgProxy.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "Simulator.h"
#include "DlgProxy.h"
#include "SimulatorDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSimulatorDlgAutoProxy

IMPLEMENT_DYNCREATE(CSimulatorDlgAutoProxy, CCmdTarget)

CSimulatorDlgAutoProxy::CSimulatorDlgAutoProxy()
{
	EnableAutomation();

	// 자동화 개체가 활성화되어 있는 동안 계속 애플리케이션을 실행하기 위해
	//	생성자에서 AfxOleLockApp를 호출합니다.
	AfxOleLockApp();

	// 애플리케이션의 주 창 포인터를 통해 대화 상자에 대한
	//  액세스를 가져옵니다.  프록시의 내부 포인터를 설정하여
	//  대화 상자를 가리키고 대화 상자의 후방 포인터를 이 프록시로
	//  설정합니다.
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CSimulatorDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CSimulatorDlg)))
		{
			m_pDialog = reinterpret_cast<CSimulatorDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CSimulatorDlgAutoProxy::~CSimulatorDlgAutoProxy()
{
	// 모든 개체가 OLE 자동화로 만들어졌을 때 애플리케이션을 종료하기 위해
	// 	소멸자가 AfxOleUnlockApp를 호출합니다.
	//  이러한 호출로 주 대화 상자가 삭제될 수 있습니다.
	if (m_pDialog != nullptr)
		m_pDialog->m_pAutoProxy = nullptr;
	AfxOleUnlockApp();
}

void CSimulatorDlgAutoProxy::OnFinalRelease()
{
	// 자동화 개체에 대한 마지막 참조가 해제되면
	// OnFinalRelease가 호출됩니다.  기본 클래스에서 자동으로 개체를 삭제합니다.
	// 기본 클래스를 호출하기 전에 개체에 필요한 추가 정리 작업을
	// 추가하세요.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CSimulatorDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSimulatorDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 참고: IID_ISimulator에 대한 지원을 추가하여
//  VBA에서 형식 안전 바인딩을 지원합니다.
//  이 IID는 .IDL 파일에 있는 dispinterface의 GUID와 일치해야 합니다.

// {7a36e8b6-03f2-4059-9cc1-3e7ff6a912c3}
static const IID IID_ISimulator =
{0x7a36e8b6,0x03f2,0x4059,{0x9c,0xc1,0x3e,0x7f,0xf6,0xa9,0x12,0xc3}};

BEGIN_INTERFACE_MAP(CSimulatorDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CSimulatorDlgAutoProxy, IID_ISimulator, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 매크로가 이 프로젝트의 pch.h에 정의됩니다.
// {db374100-169a-4389-b9b4-c6ab5e02548b}
IMPLEMENT_OLECREATE2(CSimulatorDlgAutoProxy, "Simulator.Application", 0xdb374100,0x169a,0x4389,0xb9,0xb4,0xc6,0xab,0x5e,0x02,0x54,0x8b)


// CSimulatorDlgAutoProxy 메시지 처리기
