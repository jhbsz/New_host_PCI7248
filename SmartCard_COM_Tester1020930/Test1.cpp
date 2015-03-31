// Test1.cpp : implementation file
//

#include "stdafx.h"
#include "9525COMAP.h"
#include "Test1.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Test1 dialog


Test1::Test1(CWnd* pParent /*=NULL*/)
	: CDialog(Test1::IDD, pParent)
{
	//{{AFX_DATA_INIT(Test1)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void Test1::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(Test1)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(Test1, CDialog)
	//{{AFX_MSG_MAP(Test1)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Test1 message handlers
