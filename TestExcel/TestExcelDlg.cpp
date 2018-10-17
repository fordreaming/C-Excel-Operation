// TestExcelDlg.cpp : implementation file
//

#include "stdafx.h"
#include "excel.h"
#include "TestExcel.h"
#include "TestExcelDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();
	
	// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA
	
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	
	// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
//{{AFX_MSG_MAP(CAboutDlg)
// No message handlers
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestExcelDlg dialog

CTestExcelDlg::CTestExcelDlg(CWnd* pParent /*=NULL*/)
: CDialog(CTestExcelDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTestExcelDlg)
	// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CTestExcelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTestExcelDlg)
	// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTestExcelDlg, CDialog)
//{{AFX_MSG_MAP(CTestExcelDlg)
ON_WM_SYSCOMMAND()
ON_WM_PAINT()
ON_WM_QUERYDRAGICON()
ON_BN_CLICKED(IDRUN, OnRun)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestExcelDlg message handlers

BOOL CTestExcelDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	
	// Add "About..." menu item to system menu.
	
	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);
	
	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}
	
	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CTestExcelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CTestExcelDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting
		
		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);
		
		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;
		
		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CTestExcelDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CTestExcelDlg::OnRun() 
{
	// Reference: http://msdn.microsoft.com/en-us/library/ms262200%28v=Office.11%29.aspx
	// TODO: Add your control notification handler code here
	try {
		// 先创建一个_Application类，用_Application来创建一个Excel应用程序接口
		_Application app;
		// 工作薄集合
		Workbooks books;
		// 工作薄
		_Workbook book;
		// 工作表集合
		Worksheets sheets;
		// 工作表
		_Worksheet sheet;
		// 图表
		_Chart chart;
		// 单元格区域
		Range range;
		Font font;
		Range cols;
		ChartObjects chartobjects;
		Charts charts;
		LPDISPATCH lpDisp;
		// Common OLE variants. These are easy variants to use for
		// calling arguments.
		COleVariant	covTrue((short)TRUE), covFalse((short)FALSE), 
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		
		// Start Excel and get the Application object.
		if(!app.CreateDispatch("Excel.Application"))
		{
			AfxMessageBox("Couldn't start Excel and get an application 0bject");
			return;
		}
		app.SetVisible(TRUE);

		books=app.GetWorkbooks();
		book=books.Open("D:\\Result",	// Filename
			COleVariant((short)0),		// Donot update link
			covFalse,					// Not Read-only
			COleVariant((short)3),		// Format: Delim is space
			covOptional,				// Password
			covOptional,				// WritePassword
			covTrue,					// 不显示建议只读消息
			covOptional,				//	Current OS
			covOptional,				// Delimeter
			covTrue,					//	Editable
			covFalse,					// Not Notify
			covOptional,				// Converter
			covOptional, 
			covTrue,					// Local 
			covOptional); // .Add(covOptional);
		sheets=book.GetSheets();
		sheet=sheets.GetItem(COleVariant((short)1));
		
		range = sheet.GetRange(COleVariant("B1"), COleVariant("B1"));
		range.SetValue2(COleVariant("阻塞率"));
		range = sheet.GetRange(COleVariant("C1"), COleVariant("C1"));
		range.SetValue2(COleVariant("利用率"));
		range = sheet.GetRange(COleVariant("D1"), COleVariant("D1"));
		range.SetValue2(COleVariant("消耗时间"));
		range = sheet.GetRange(COleVariant("A2"), COleVariant("A2"));
		range.SetValue2(COleVariant("未分域"));
		range = sheet.GetRange(COleVariant("A3"), COleVariant("A3"));
		range.SetValue2(COleVariant("分域后"));
		
		range = sheet.GetRange(COleVariant("B2"), COleVariant("B2"));
		range.SetValue2(COleVariant("0.2"));
		range = sheet.GetRange(COleVariant("C2"), COleVariant("C2"));
		range.SetValue2(COleVariant("0.5"));
		range = sheet.GetRange(COleVariant("D2"), COleVariant("D2"));
		range.SetValue2(COleVariant("10"));
		
		
		range = sheet.GetRange(COleVariant("B3"), COleVariant("B3"));
		range.SetValue2(COleVariant("0.3"));
		range = sheet.GetRange(COleVariant("C3"), COleVariant("C3"));
		range.SetValue2(COleVariant("0.8"));
		range = sheet.GetRange(COleVariant("D3"), COleVariant("D3"));
		range.SetValue2(COleVariant("9"));
	
		
		// The cells are populated. To start the chart,
		// declare some long variables and site the chart.
		long left, top, width, height;
		left = 200;
		top = 50;
		width = 350;
		height = 250;
		
		lpDisp = sheet.ChartObjects(covOptional);
		ASSERT(lpDisp);
		chartobjects.AttachDispatch(lpDisp); // Attach the lpDisp pointer
		// for ChartObjects to the chartobjects
		// object.
		if (chartobjects.GetCount() != 0) //当excel中存在原有图表时，删除之
		{
			chartobjects.Delete();
		}
		ChartObject chartobject = chartobjects.Add(left, top, width, height); 

		chart.AttachDispatch(chartobject.GetChart()); // GetChart() returns
		// LPDISPATCH, and this attaches 
		// it to your chart object.
		
		lpDisp = sheet.GetRange(COleVariant("A1"), COleVariant("D3"));
		// The range containing the data to be charted.
		ASSERT(lpDisp);
		range.AttachDispatch(lpDisp);
		
		VARIANT var; // ChartWizard needs a Variant for the Source range.
		var.vt = VT_DISPATCH; // .vt is the usable member of the tagVARIANT
		// Struct. Its value is a union of options.
		var.pdispVal = lpDisp; // Assign IDispatch pointer
		// of the Source range to var.
		
		chart.ChartWizard(var,       // Source.
			COleVariant((short)11),  // Gallery: 3d Column.
			covOptional,             // Format, use default.
			COleVariant((short)1),   // PlotBy: xlRows.
			COleVariant((short)1),   // CategoryLabels. 第一行是分类标签
			COleVariant((short)1),   // SeriesLabels. 第一列是系列标签
			COleVariant((short)TRUE), // HasLegend.
			COleVariant("仿真结果对比图"),  // Title.
			COleVariant("结果类别"),    // CategoryTitle.
			COleVariant("数值"),  // ValueTitles.
			covOptional              // ExtraTitle.
			);



	//***************************************************************************

		chartobject = chartobjects.Add(left+1000, top, width, height); 

		chart.AttachDispatch(chartobject.GetChart()); // GetChart() returns
		// LPDISPATCH, and this attaches 
		// it to your chart object.

		lpDisp = sheet.GetRange(COleVariant("A1"), COleVariant("D3"));
		// The range containing the data to be charted.
		ASSERT(lpDisp);
		range.AttachDispatch(lpDisp);

		//VARIANT var; // ChartWizard needs a Variant for the Source range.
		var.vt = VT_DISPATCH; // .vt is the usable member of the tagVARIANT
		// Struct. Its value is a union of options.
		var.pdispVal = lpDisp; // Assign IDispatch pointer
		// of the Source range to var.

		chart.ChartWizard(var,       // Source.
			COleVariant((short)11),  // Gallery: 3d Column.
			covOptional,             // Format, use default.
			COleVariant((short)1),   // PlotBy: xlRows.
			COleVariant((short)1),   // CategoryLabels. 第一行是分类标签
			COleVariant((short)1),   // SeriesLabels. 第一列是系列标签
			COleVariant((short)TRUE), // HasLegend.
			COleVariant("仿真结果对比图"),  // Title.
			COleVariant("结果类别"),    // CategoryTitle.
			COleVariant("数值"),  // ValueTitles.
			covOptional              // ExtraTitle.
			);
		
      }  // End of processing logic.
	  
      catch(COleException *e)
      {
		  char buf[1024];
		  
		  sprintf(buf, "COleException. SCODE: %08lx.", (long)e->m_sc);
		  ::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
      }
	  
      catch(COleDispatchException *e)
      {
		  char buf[1024];
		  sprintf(buf,
			  "COleDispatchException. SCODE: %08lx, Description: \"%s\".",
			  (long)e->m_wCode,
			  (LPSTR)e->m_strDescription.GetBuffer(1024));
		  ::MessageBox(NULL, buf, "COleDispatchException",
			  MB_SETFOREGROUND | MB_OK);
      }
	  
      catch(...)
      {
		  ::MessageBox(NULL, "General Exception caught.", "Catch-All",
			  MB_SETFOREGROUND | MB_OK);
      }
}
