
// SimulatorDlg.h: 헤더 파일
//

#pragma once

#include "GlobalEnv.h"

class CSimulatorDlgAutoProxy;


// CSimulatorDlg 대화 상자
class CSimulatorDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CSimulatorDlg);
	friend class CSimulatorDlgAutoProxy;

// 생성입니다.
public:
	CSimulatorDlg(CWnd* pParent = nullptr);	// 표준 생성자입니다.
	virtual ~CSimulatorDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SIMULATOR_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


// 구현입니다.
protected:
	CSimulatorDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedCreateProject();

//song
public:
	void GetMainParameters(ALL_ACT_TYPE* act, ALL_ACTIVITY_PATTERN* pattern);
	void MakeProjectAndRun(CString strFileName, CString strInSheetName, CString strOutSheetName);
	void OnlyRun(CString strFileName, CString strInSheetName, CString strOutSheetName);
	void MakeResult(CString strFileName, CString  strResultSheetName, CString  strOutSheetName, CString strInSheetName, int num, BOOL isDelete = TRUE);
	GLOBAL_ENV* m_pGlobalEnv;
	afx_msg void OnBnClickedSimulation();
	CProgressCtrl* m_Progress;



	//song
private:	
	CString m_strFileName;
	int m_ProblemCnt = 100;
	CString m_strComment;

	void MakePath();
	BOOL CreateDirectoryRecursively(const CString& strPath);

public:
	afx_msg void OnBnClickedClearXl();
	afx_msg void OnBnClickedOnlyRun();	
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedChangeFolder();
	afx_msg void OnEnChangeEdit13();
};
