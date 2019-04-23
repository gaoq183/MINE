// CAPPTreeView.cpp : implementation file
//

#include "stdafx.h"
#include "kbmanage.h"
#include "CAPPTreeView.h"

#include "UpdtClsPS.h"
#include "kbcomm.h"
#include "UpdtClsNameDlg.h"
#include "currentuser.h"
#include "mainfrm.h"
#include "DefClsPS.h"
#include "KDX.h"
#include "Table.h"
#include "KbclassEx.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern CKbmApp theApp;
CKBClass cCurCls;
/////////////////////////////////////////////////////////////////////////////
// CCAPPTreeView

IMPLEMENT_DYNCREATE(CCAPPTreeView, CTreeView)

CCAPPTreeView::CCAPPTreeView()
{
}

CCAPPTreeView::~CCAPPTreeView()
{
}


BEGIN_MESSAGE_MAP(CCAPPTreeView, CTreeView)
//{{AFX_MSG_MAP(CCAPPTreeView)
ON_COMMAND(ID_UpdateCls, OnRightChange)
ON_COMMAND(ID_AddCls, OnEditNewItem)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCAPPTreeView drawing

void CCAPPTreeView::OnDraw(CDC* pDC)
{
	CDocument* pDoc = GetDocument();
	// TODO: add draw code here
}

/////////////////////////////////////////////////////////////////////////////
// CCAPPTreeView diagnostics

#ifdef _DEBUG
void CCAPPTreeView::AssertValid() const
{
	CTreeView::AssertValid();
}

void CCAPPTreeView::Dump(CDumpContext& dc) const
{
	CTreeView::Dump(dc);
}

CKbmDoc* CCAPPTreeView::GetDocument() // non-debug version is inline
{
	return (CKbmDoc*)((CMainFrame*)AfxGetMainWnd())->MDIGetActive()->GetActiveDocument();

	//	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CKbmDoc)));
	//	return (CKbmDoc*)m_pDocument;
}

#else
CKbmDoc* CCAPPTreeView::GetDocument() // non-debug version is inline
{
	return (CKbmDoc*)((CMainFrame*)AfxGetMainWnd())->MDIGetActive()->GetActiveDocument();
	//	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CKbmDoc)));
	//	return (CKbmDoc*)m_pDocument;
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CCAPPTreeView message handlers

int CCAPPTreeView::ListSubCls(CKBClass * pCls)
{
	CObList cClsList;
	CAPP_RCODE rc=pCls->KB_list_direct_SubClsIdNameById(&cClsList);
	KB_free_cAttrList(&pCls->m_cAttrList);
	if(rc==CAPP_ERROR)
		return CAPP_ERROR;
	else if(rc==CAPP_NO_DATA_FOUND)
		return CAPP_NO_DATA_FOUND;
	
	POSITION pos = cClsList.GetHeadPosition();
	while (pos != NULL)
	{
		CKBClass * pClsBuf = (CKBClass*)cClsList.GetNext(pos);
		////////////gp modified
		//判断是否需要截断典型工艺
		CKBClassEx cKbClsEx;
		cKbClsEx.m_idCls=pClsBuf->m_idCls;
		if(cKbClsEx.Get_ParentClsIDList_by_ID()==CAPP_ERROR){
			while(!cClsList.IsEmpty())
				delete (CKBClass*)cClsList.RemoveHead();
			return CAPP_ERROR;
		}
		if(cKbClsEx.m_IDList.GetCount()>4&&(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(1))==100001
			&&(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(2))==100018){
			CKBClass cTempKBClass;
			cTempKBClass.m_idCls=(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(4));
			char  strTempItem[128];
			cTempKBClass.KB_get_cls_name();
			strcpy(strTempItem,cTempKBClass.m_sClsName);
			strTempItem[4]='\0';
			if(strcmp(strTempItem,"典型")==0)
				continue;
		}

		CString strItem;
		strItem=pClsBuf->m_sClsName;
		tClassItem=tClassItem.InsertAfter(strItem, TVI_LAST,IID_CLASS);
		tClassItem.EnsureVisible();
		tClassItem.SetData((DWORD)pClsBuf->m_idCls);
		m_bNoNotifications = TRUE;
		if(ListSubCls(pClsBuf)!=CAPP_SUCCESS){
			tClassItem=tClassItem.GetParent();
			continue;
		}
		//	  CTreeCtrlEx& ctlTree = (CTreeCtrlEx&) GetTreeCtrl();
		//     ctlTree.CollpseBranch(tClassItem);
		//释放类对象
		tClassItem=tClassItem.GetParent();
	}   
	while(!cClsList.IsEmpty())
		delete (CKBClass*)cClsList.RemoveHead();

	tClassItem.EnsureVisible();
	return CAPP_SUCCESS;
}


int CCAPPTreeView::ListCls(CKBClass * pCls)
{
	CObList cClsList;
	CAPP_RCODE rc=pCls->KB_list_direct_SubClsIdNameById(&cClsList);
	KB_free_cAttrList(&pCls->m_cAttrList);
	if(rc==CAPP_ERROR)
		return CAPP_ERROR;
	else if(rc==CAPP_NO_DATA_FOUND)
		return CAPP_NO_DATA_FOUND;
	
	POSITION pos = cClsList.GetHeadPosition();
	while (pos != NULL)
	{
		CKBClass * pClsBuf = (CKBClass*)cClsList.GetNext(pos);
		///////////gp modified
		//判断是否需要截断典型工艺
		CKBClassEx cKbClsEx;
		cKbClsEx.m_idCls=pClsBuf->m_idCls;
		if(cKbClsEx.Get_ParentClsIDList_by_ID()==CAPP_ERROR){
			while(!cClsList.IsEmpty())
				delete (CKBClass*)cClsList.RemoveHead();
			return CAPP_ERROR;
		}
		if(cKbClsEx.m_IDList.GetCount()>4&&(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(1))==100001
			&&(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(2))==100018){
			CKBClass cTempKBClass;
			cTempKBClass.m_idCls=(OBJECT_ID)cKbClsEx.m_IDList.GetAt(cKbClsEx.m_IDList.FindIndex(4));
			char  strTempItem[128];
			cTempKBClass.KB_get_cls_name();
			strcpy(strTempItem,cTempKBClass.m_sClsName);
			strTempItem[4]='\0';
			if(strcmp(strTempItem,"典型")==0)
				continue;
		}

		CString strItem;
		strItem=pClsBuf->m_sClsName;
		tClassItem=tClassItem.InsertAfter(strItem, TVI_LAST,IID_CLASS);
		tClassItem.EnsureVisible();
		tClassItem.SetData((DWORD)pClsBuf->m_idCls);
		m_bNoNotifications = TRUE;
		tClassItem = tClassItem.GetParent();
	}   
	while(!cClsList.IsEmpty())
		delete (CKBClass*)cClsList.RemoveHead();

	tClassItem.EnsureVisible();
	return CAPP_SUCCESS;
}

void CCAPPTreeView::OnRightChange() 
{
	CKBClass cKBCls;

   cKBCls.m_idCls=(OBJECT_ID)m_ItemSel.GetData();
   if(cKBCls.KB_get_cls_withCFF()==CAPP_ERROR)
   {
      MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
      return;
   }
   char sOldClsName[64];
   strcpy(sOldClsName,cKBCls.m_sClsName);

	//==========liuhy modify begin 2003.06.25==================================
	//封掉原来的帮助按钮
	CUpdtClsPS UpdtClsPS;
	UpdtClsPS.m_psh.dwFlags &= ~(PSH_HASHELP);
	//==========liuhy modify end 2003.06.25====================================

	UpdtClsPS.m_cAttrPage.m_idCls=(OBJECT_ID)m_ItemSel.GetData();
	UpdtClsPS.m_cClsInfoPage.m_idCls=UpdtClsPS.m_cAttrPage.m_idCls;
	CKbmDoc* pDoc=GetDocument();
	UpdtClsPS.m_cAttrPage.m_pClsMap=&pDoc->m_cClsMap;
	UpdtClsPS.m_cAttrPage.m_idCls=(OBJECT_ID)m_ItemSel.GetData();
	UpdtClsPS.m_cClsGraphPage.m_csGraphFileName=cKBCls.m_sReserved5;
	UpdtClsPS.m_cClsGraphPage.m_DocFile.m_idHaving=UpdtClsPS.m_cAttrPage.m_idCls;
	UpdtClsPS.m_cClsInfoPage.m_bCFF = cKBCls.m_bCFF;
	UpdtClsPS.m_cClsInfoPage.m_sSupClsName=(m_ItemSel.GetParent()).GetText();
	UpdtClsPS.m_cClsInfoPage.m_nClsType=cKBCls.m_nClsType;
	UpdtClsPS.m_cClsInfoPage.m_sClsName=m_ItemSel.GetText();
	UpdtClsPS.m_cClsInfoPage.m_nLinkDataType = cKBCls.m_nLinkDataType;
   switch(cKBCls.m_nLinkDataType)
   {
   case KBDATA:
      UpdtClsPS.m_cClsInfoPage.m_nType = 0;
      break;
   case EDBDATA:
      UpdtClsPS.m_cClsInfoPage.m_sDataSource = cKBCls.m_sDataBaseName;
      UpdtClsPS.m_cClsInfoPage.m_sTableName = cKBCls.m_sTableName;
      UpdtClsPS.m_cClsInfoPage.m_sQueryStr = cKBCls.m_sQueryStr;
      UpdtClsPS.m_cClsInfoPage.m_nType = 1;
      if(cKBCls.m_nReserved1 == LINKDB)
         UpdtClsPS.m_cClsInfoPage.m_nLinkDB = 0;
      else
         UpdtClsPS.m_cClsInfoPage.m_nLinkDB = 1;
      break;
   case IDBDATA:
      UpdtClsPS.m_cClsInfoPage.m_sTableName = cKBCls.m_sTableName;
      UpdtClsPS.m_cClsInfoPage.m_sQueryStr = cKBCls.m_sQueryStr;
      UpdtClsPS.m_cClsInfoPage.m_nType = 2;
      if(cKBCls.m_nReserved1 == LINKDB)
         UpdtClsPS.m_cClsInfoPage.m_nLinkDB = 0;
      else
         UpdtClsPS.m_cClsInfoPage.m_nLinkDB = 1;
      break;
   }

   if(UpdtClsPS.DoModal()==IDOK)
   {
		theApp.m_bDbClickAttrorInst=FALSE;
		CKBClass *pcKBCls=new CKBClass;
      pcKBCls->m_bIsReserved=TRUE;
      pcKBCls->m_idSuperCls=(OBJECT_ID)(m_ItemSel.GetParent()).GetData();
      pcKBCls->m_idCls=(OBJECT_ID)m_ItemSel.GetData();
      strcpy(pcKBCls->m_sClsName,UpdtClsPS.m_cClsInfoPage.m_sClsName);
      strcpy(pcKBCls->m_sReserved5,(const char*)UpdtClsPS.m_cClsGraphPage.m_csGraphFileName);
      pcKBCls->m_nLinkDataType = UpdtClsPS.m_cClsInfoPage.m_nLinkDataType;
      switch(pcKBCls->m_nLinkDataType)
      {
      case EDBDATA:
         strcpy(pcKBCls->m_sDataBaseName, UpdtClsPS.m_cClsInfoPage.m_sDataSource);
         strcpy(pcKBCls->m_sTableName, UpdtClsPS.m_cClsInfoPage.m_sTableName);
         strcpy(pcKBCls->m_sQueryStr, UpdtClsPS.m_cClsInfoPage.m_sQueryStr);
         if(UpdtClsPS.m_cClsInfoPage.m_nLinkDB == 0)
            pcKBCls->m_nReserved1 = LINKDB;
         else
            pcKBCls->m_nReserved1 = NOLINKDB;   
         break;
      case IDBDATA:
         strcpy(pcKBCls->m_sTableName, UpdtClsPS.m_cClsInfoPage.m_sTableName);
         strcpy(pcKBCls->m_sQueryStr, UpdtClsPS.m_cClsInfoPage.m_sQueryStr);
         if(UpdtClsPS.m_cClsInfoPage.m_nLinkDB == 0)
            pcKBCls->m_nReserved1 = LINKDB;
         else
            pcKBCls->m_nReserved1 = NOLINKDB;   
         break;
      }
	  pcKBCls->m_bCFF = UpdtClsPS.m_cClsInfoPage.m_bCFF;

      POSITION pos=UpdtClsPS.m_cAttrPage.m_Grid.m_cAttrList.GetHeadPosition();
      while(pos!=NULL)
      {
         CKBAttr *pcAttr=(CKBAttr*)UpdtClsPS.m_cAttrPage.m_Grid.m_cAttrList.GetNext(pos);
         pcKBCls->m_cAttrList.AddTail(pcAttr);
      }
//      pcKBMCls->m_Mthd.AddTail(&DefClsPS.m_cMthdPage.m_cMthd);
      BeginWaitCursor();
      if(CAPP_parse_name_str(pcKBCls->m_sClsName,strlen(pcKBCls->m_sClsName))==CAPP_ERROR)
      {
         strcat(pcKBCls->m_sClsName,"新建类XXX");
         MessageBox("请更改类名","提示信息",MB_OK);
      }
      OBJECT_ID idTemp;
      if(KB_find_cls(pcKBCls->m_sClsName,&idTemp)==CAPP_SUCCESS)
      {
         if(idTemp==pcKBCls->m_idCls)
            goto end;
//begin:
         char sTemp[200];
         sprintf(sTemp,"类<%s>已存在，请更改类名",pcKBCls->m_sClsName);
         MessageBox(sTemp,"提示信息",MB_OK);
         CUpdtClsNameDlg cUCNDlg;
         cUCNDlg.m_csClsName=pcKBCls->m_sClsName;
         if(cUCNDlg.DoModal()==IDOK)
            strcpy(pcKBCls->m_sClsName,(const char*)cUCNDlg.m_csClsName);
         else
            strcpy(pcKBCls->m_sClsName,sOldClsName);
      }
end:
      CTime t,t1;
      t1=t.GetCurrentTime( ); // time_t defined in <time.h>
      CString csStr=t1.Format("%Y-%m-%d %H:%M:%S");

      strcpy(pcKBCls->m_sUpdtUser,gCurrentUser.m_sUserName);
      strcpy(pcKBCls->m_sUpdtTime,(const char*)csStr);
//2002.8.2.10 gp	更新类可以任意改变属性的值类型
//如果检查出属性值类型发生改变，提示后进行更新
		CObList AttrChangedList;			//值类型发生改变的属性，来自pcKBCls->m_cAttrList，不需要释放节点
		if(PPBT_CheckAttrTypeChanged(pcKBCls,AttrChangedList)!=CAPP_SUCCESS){
			MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
			delete pcKBCls;
			return;
		}
 
		if(!AttrChangedList.IsEmpty()){
			if(MessageBox("发现有属性的值类型发生改变，可能会丢失数据，是否确实需要改变？","严重警告",MB_YESNO | MB_ICONQUESTION)==IDYES){
				if(PPB_UpdateClass_Attr_Col(pcKBCls,AttrChangedList)==CAPP_ERROR){
					MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
					delete pcKBCls;
					return;
				}
			}
			else	{
				delete pcKBCls;
				return;
			}
		}
		else{
			if(pcKBCls->KB_updt_cls_withCFF()==CAPP_ERROR)
			{
				MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
				delete pcKBCls;
				return;
			}
		}
//2002.9.7.11 gp插入图形文件
		UpdtClsPS.m_cClsGraphPage.m_DocFile.m_idHaving=pcKBCls->m_idCls;
		if(UpdtClsPS.m_cClsGraphPage.MakeFile()==CAPP_ERROR){					//包括新增、删除、替换库中文件
         MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
         return;
		}

      EndWaitCursor();

      char sClsID[18];
      sprintf(sClsID,"%I64d",pcKBCls->m_idCls);
      pDoc->m_cClsMap.SetAt(sClsID,pcKBCls);
//      m_pcCurCls=pcKBCls;
		     
      m_ItemSel.SetText(pcKBCls->m_sClsName);
		m_ItemSel.EnsureVisible();
	  
      m_ItemSel.Select();
		//tClass.SortChildren();
		m_bNoNotifications = TRUE;
//		if(tClassItem.EditLabel() == NULL)
//		  return;
	}
}

void CCAPPTreeView::OnEditNewItem()
{
	UINT nImageID = m_ItemSel.GetImageID();

	//==========liuhy modify begin 2003.06.25==================================
	//封掉原来的帮助按钮	
	CDefClsPS DefClsPS;
	DefClsPS.m_psh.dwFlags &= ~(PSH_HASHELP);
	//==========liuhy modify end 2003.06.25====================================
	DefClsPS.m_cClsInfoPage.m_sSupClsName=m_ItemSel.GetText();
	char sNewClsName[32];
	sprintf(sNewClsName,"新建类%d",m_nClsSNo);
	m_nClsSNo++;
	DefClsPS.m_cClsInfoPage.m_sClsName=sNewClsName;
	DefClsPS.m_cAttrPage.m_idSupCls=(OBJECT_ID)m_ItemSel.GetData();
	DefClsPS.m_cClsInfoPage.m_idSupCls=DefClsPS.m_cAttrPage.m_idSupCls;
	if(DefClsPS.DoModal()==IDOK)
	{
		theApp.m_bDbClickAttrorInst=FALSE;
		CKbmDoc* pDoc=GetDocument();
		CKBClass *pcKBCls=new CKBClass;
		
		CTime t,t1;
		t1=t.GetCurrentTime( ); // time_t defined in <time.h>
		CString csStr=t1.Format("%Y-%m-%d %H:%M:%S");
		
		strcpy(pcKBCls->m_sCreateUser,gCurrentUser.m_sUserName);
		strcpy(pcKBCls->m_sCreateTime,(const char*)csStr);
		
		pcKBCls->m_bIsReserved=TRUE;
		pcKBCls->m_idSuperCls=(OBJECT_ID)m_ItemSel.GetData();//DefClsPS.m_cClsInfoPage.m_idSupCls;
		strcpy(pcKBCls->m_sClsName,DefClsPS.m_cClsInfoPage.m_sClsName);
		pcKBCls->m_nClsType=DefClsPS.m_cClsInfoPage.m_nClsType;
		strcpy(pcKBCls->m_sReserved5,(const char*)DefClsPS.m_cClsGraphPage.m_csGraphFileName);
		pcKBCls->m_nLinkDataType = DefClsPS.m_cClsInfoPage.m_nLinkDataType;  
		switch(pcKBCls->m_nLinkDataType)
		{
		case EDBDATA:
			strcpy(pcKBCls->m_sDataBaseName, DefClsPS.m_cClsInfoPage.m_sDataSource);
			strcpy(pcKBCls->m_sTableName, DefClsPS.m_cClsInfoPage.m_sTableName);
			strcpy(pcKBCls->m_sQueryStr, DefClsPS.m_cClsInfoPage.m_sQueryStr);
			if(DefClsPS.m_cClsInfoPage.m_nLinkDB == 0)
				pcKBCls->m_nReserved1 = LINKDB;
			else
				pcKBCls->m_nReserved1 = NOLINKDB;
			break;
		case IDBDATA:
			strcpy(pcKBCls->m_sTableName, DefClsPS.m_cClsInfoPage.m_sTableName);
			strcpy(pcKBCls->m_sQueryStr, DefClsPS.m_cClsInfoPage.m_sQueryStr);
			if(DefClsPS.m_cClsInfoPage.m_nLinkDB == 0)
				pcKBCls->m_nReserved1 = LINKDB;
			else
				pcKBCls->m_nReserved1 = NOLINKDB;
			break;
		}
		pcKBCls->m_cAttrList.RemoveAll();
		POSITION pos=DefClsPS.m_cAttrPage.m_Grid.m_cAttrList.GetHeadPosition();
		while(pos!=NULL)
		{
			CKBAttr *pcAttr=(CKBAttr*)DefClsPS.m_cAttrPage.m_Grid.m_cAttrList.GetNext(pos);
			pcKBCls->m_cAttrList.AddTail(pcAttr);
		}
		//      pcKBMCls->m_Mthd.AddTail(&DefClsPS.m_cMthdPage.m_cMthd);
		BeginWaitCursor();
		if(CAPP_parse_name_str(pcKBCls->m_sClsName,strlen(pcKBCls->m_sClsName))==CAPP_ERROR)
		{
			strcat(pcKBCls->m_sClsName,"新建类");
			MessageBox("请更改类名","提示信息",MB_OK);
		}
/*
		OBJECT_ID idTemp;
		if(KB_find_cls(pcKBCls->m_sClsName,&idTemp)==CAPP_SUCCESS)
		{
begin:
		char sTemp[200];
		sprintf(sTemp,"类<%s>已存在，请更改类名",pcKBCls->m_sClsName);
		MessageBox(sTemp,"提示信息",MB_OK);
		CUpdtClsNameDlg cUCNDlg;
		cUCNDlg.m_csClsName=pcKBCls->m_sClsName;
		if(cUCNDlg.DoModal()==IDOK)
			strcpy(pcKBCls->m_sClsName,(const char*)cUCNDlg.m_csClsName);
		else
			goto begin;
		}
*/		
		pcKBCls->m_bCFF = DefClsPS.m_cClsInfoPage.m_bCFF;
		if(pcKBCls->KB_add_cls_withCFF()==CAPP_ERROR)
		{
			MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
			return;
		}
//2002.9.7.11 gp插入图形文件
		DefClsPS.m_cClsGraphPage.m_DocFile.m_idHaving=pcKBCls->m_idCls;
		if(DefClsPS.m_cClsGraphPage.MakeFile()==CAPP_ERROR){					//包括新增、删除、替换库中文件
         MessageBox(sCAPPErrMsg,"提示信息",MB_OK);
         return;
		}

		EndWaitCursor();
		
		char sClsID[18];
		sprintf(sClsID,"%I64d",pcKBCls->m_idCls);
		pDoc->m_cClsMap.SetAt(sClsID,pcKBCls);
		//      m_pcCurCls=pcKBCls;
		
		tClassItem=m_ItemSel.InsertAfter(pcKBCls->m_sClsName, 
			TVI_LAST,IID_CLASS);
		tClassItem.EnsureVisible();
		
		m_ItemSel.Select();
		//tClass.SortChildren();
		m_bNoNotifications = TRUE;
		m_ItemSel = tClassItem;
		m_ItemSel.SetData((DWORD)pcKBCls->m_idCls);
	}
}



void CCAPPTreeView::DeleteItem(CTreeCursor& itemDelete)
{
	itemDelete = m_ItemSel;
	CString sStr;
	OBJECT_ID idCls=(OBJECT_ID)itemDelete.GetData();
	if(idCls>=100000&&idCls<=100049)
	{
		MessageBox("该类为支持系统运行所必需的类,不能删除",_T("提示信息"),MB_OK);
		return;
	}
	
	//判断当前类有没有被其他已经引用，如果引用了，不允许删除 begin 
	char strbuf[LONGSTRINGLEN+1];
	char idbuf[4][18];
	RETCODE rc , rc1;
	SDWORD cbIdAttribute[23]; 
		
	sprintf(strbuf,"SELECT idObjCls FROM KBT_ATTR WHERE idObjCls ='%I64d'",idCls);
	rc = SQLExecDirect(ppb_hstmt,(unsigned char*)strbuf,SQL_NTS);	
	if(RETCODE_IS_FAILURE(rc))
	{
		rc1=SQLError(SQL_NULL_HENV,SQL_NULL_HDBC,ppb_hstmt,szSqlState,&dwNativeError,
			szErrorMsg,sizeof(szErrorMsg),&wErrorMsg);     
		strcpy(sCAPPErrMsg ,(const char *)szErrorMsg);
		AfxMessageBox(sCAPPErrMsg);
		rc=ClearPPBStmt(&ppb_hstmt);             
		return;
	}
	rc = SQLBindCol(ppb_hstmt,1,SQL_C_CHAR,idbuf[0],18,&cbIdAttribute[0]);
	if(RETCODE_IS_FAILURE(rc))
	{
		rc1=SQLError(SQL_NULL_HENV,SQL_NULL_HDBC,ppb_hstmt,szSqlState,&dwNativeError,
			szErrorMsg,sizeof(szErrorMsg),&wErrorMsg);     
		strcpy(sCAPPErrMsg ,(const char *)szErrorMsg);
		AfxMessageBox(sCAPPErrMsg);
		rc=ClearPPBStmt(&ppb_hstmt);             
		return;
	}
	rc=SQLFetch(ppb_hstmt);
	if(rc != SQL_NO_DATA_FOUND)
	{
		ClearPPBStmt(&ppb_hstmt);
		AfxMessageBox("该类已经被其他类所引用，您不能删除!");
		return;
	}
	ClearPPBStmt(&ppb_hstmt);
	//判断当前类有没有被其他已经引用，如果引用了，不允许删除 end
	
	switch(itemDelete.GetImageID())
	{
	case IID_CLASS:
		{
			int retCode;
			sStr="你确实要删除类 <"+itemDelete.GetText()+"> 及其子类吗?";
			retCode = MessageBox(sStr,_T("提示信息"),MB_YESNO);
			if (retCode == IDYES)
			{
				sStr="删除类 <"+itemDelete.GetText()+"> 后,该类及其子类信息将全部被删除，要删除吗?";
				if(MessageBox(sStr,_T("提示信息"),MB_YESNO)==IDYES)
				{
					CKBClass cCls;
					cCls.m_idCls=(OBJECT_ID)itemDelete.GetData();
					BeginWaitCursor();
					if(cCls.KB_del_cls()==CAPP_ERROR)
					{
						MessageBox("删除类错误",_T("提示信息"),MB_OK);
						return;
					}
					
					if((OBJECT_ID)itemDelete.GetData()==cCurCls.m_idCls)
					{
						CMainFrame *pMainFram=(CMainFrame*)AfxGetApp()->m_pMainWnd;
						if(pMainFram)
						{
							pMainFram->SendMessage(WM_COMMAND,ID_FILE_CLOSE);
							pMainFram->SendMessage(WM_COMMAND,ID_FILE_NEW);
						}
					}
					
					CKbmDoc* pDoc=GetDocument();
					char sClsID[18];
					sprintf(sClsID,"%I64d",cCls.m_idCls);
					pDoc->m_cClsMap.RemoveKey(sClsID);
					KB_free_cAttrList(&cCls.m_cAttrList);
					itemDelete.Delete();
					EndWaitCursor();
				}
			}
			break;
		}
	}
}


void CCAPPTreeView::OnInstToIndpdtTbl(OBJECT_ID idCls) 
{
	CodbcDB		cDB;
	CAPP_RCODE  rc;
	rc = CKDX::KBIO_InstToIndpdtTbl(idCls, &cDB);
	
	if(rc==CAPP_CANCEL)	return;
	
	if(rc != CAPP_SUCCESS)
	{
		if(rc!=CAPP_NO_DATA_FOUND)
			MessageBox("操作失败，事务将回退。","出错信息",MB_OK|MB_ICONSTOP);
		Rollback(cDB.m_hdbc);
		goto err;
	} 
	
	if( !Commit(cDB.m_hdbc) )
	{
		MessageBox("操作失败，事务将回退。","出错信息",MB_OK|MB_ICONSTOP);
		Rollback(cDB.m_hdbc);
		goto err;
	}
	
	MessageBox("操作完成。","类的实例复制为独立表", MB_OK|MB_ICONEXCLAMATION);
err:
	if( cDB.m_hdbc == hdbc )
		cDB.m_hdbc = 0;
}
/*
void CCAPPTreeView::OnImportExtern() 
{
	OBJECT_ID idCls = m_ItemSel.GetData();
	CString sClsName = m_ItemSel.GetText();
	char sClsName1[255];
	strcpy(sClsName1,sClsName);
	BeginWaitCursor();
	//调用外部ExcelToMDB程序，将Excel文件转换成ACCESS数据库
	//动态配置数据源
	// Initialize the STARTUPINFO structure.
	CString	filename="";
	filename.Format("%s\\%s.mdb",jCAPPEnv.sBufPath,sClsName1);
	unlink(filename);
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	memset(&si, 0, sizeof(si));
	si.cb = sizeof(si);
	
	CreateProcess(
		NULL,
		"ExcelToMDB",
		NULL,	// pointer to process security attributes 
		NULL,	// pointer to thread security attributes 
		FALSE,	// handle inheritance flag 
		0,		// creation flags 
		NULL,	// pointer to new environment block 
		NULL,	// pointer to current directory name 
		&si,	// pointer to STARTUPINFO 
		&pi		// pointer to PROCESS_INFORMATION  
		);
	
	HANDLE  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pi.dwProcessId);
	DWORD ExitEvent = WaitForSingleObject(hProcess, 99999999999999999);   
	//Find the *.mdb exit
	CFileFind finder;   
	if(finder.FindFile(filename)==TRUE)
	{
		CKDX cKDX;
		if(cKDX.KBIO_IndpdtTblToInst(sClsName1,idCls) == CAPP_ERROR)
		{
			AfxMessageBox("导入外部EXCEL文件错误!");
			EndWaitCursor();
			return;
		}
		//Repaint kbmview
		CKbmDoc* pDoc = GetDocument();
		ASSERT_VALID(pDoc);
		if(!pDoc->RepaintView(m_ItemSel.GetText(), m_ItemSel.GetData())) 	  
		{
			EndWaitCursor();
			return;
		}
		unlink(filename);
	}
	finder.Close();
	EndWaitCursor();
}
*/

void CCAPPTreeView::OnImportExtern() 
{
 	OBJECT_ID idCls = m_ItemSel.GetData();
   CString sClsName = m_ItemSel.GetText();
   char sClsName1[255];
   strcpy(sClsName1,sClsName);
   BeginWaitCursor();
   //调用外部ExcelToMDB程序，将Excel文件转换成ACCESS数据库
   //动态配置数据源
	// Initialize the STARTUPINFO structure.
	CString	filename="";
   filename.Format("%s\\%s.mdb",jCAPPEnv.sBufPath,sClsName1);
   unlink(filename);
   STARTUPINFO si;
	PROCESS_INFORMATION pi;
	memset(&si, 0, sizeof(si));
	si.cb = sizeof(si);

	CreateProcess(
		NULL,
		"ExcelToMDB",
		NULL,	// pointer to process security attributes 
		NULL,	// pointer to thread security attributes 
		FALSE,	// handle inheritance flag 
		0,		// creation flags 
		NULL,	// pointer to new environment block 
		NULL,	// pointer to current directory name 
		&si,	// pointer to STARTUPINFO 
		&pi		// pointer to PROCESS_INFORMATION  
	   );

   HANDLE  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pi.dwProcessId);
   DWORD ExitEvent = WaitForSingleObject(hProcess, 99999999999999999);   
   //Find the *.mdb exit
	CFileFind finder;   
   if(finder.FindFile(filename)==TRUE)
   {
		CKDX cKDX(0);
      if(cKDX.KBIO_IndpdtTblToInst(sClsName1,idCls) == CAPP_ERROR)
      {
	      AfxMessageBox("导入外部EXCEL文件错误!");
         EndWaitCursor();
         return;
      }
      //Repaint kbmview
      CKbmDoc* pDoc = GetDocument();
      ASSERT_VALID(pDoc);
      if(!pDoc->RepaintView(m_ItemSel.GetText(), m_ItemSel.GetData())) 	  
      {
         EndWaitCursor();
         return;
      }
      unlink(filename);
   }
	finder.Close();
   EndWaitCursor();
}

CAPP_RCODE CCAPPTreeView::PPBT_CheckAttrTypeChanged(CKBClass* pKBCls,CObList& AttrChangedList){
	CKBClass OriKBCls;
	OriKBCls.m_idCls=pKBCls->m_idCls;
	if(OriKBCls.KB_get_attr_list()!=CAPP_SUCCESS)
		return CAPP_ERROR;
	POSITION posOri=OriKBCls.m_cAttrList.GetHeadPosition();
	CKBAttr *pAttrOri,* pAttrTarg;
//要求定义的属性名称不同
	while(posOri){
		pAttrOri= (CKBAttr *)OriKBCls.m_cAttrList.GetNext(posOri);
		POSITION posTarg=pKBCls->m_cAttrList.GetHeadPosition();
		while(posTarg){
			pAttrTarg=(CKBAttr *)pKBCls->m_cAttrList.GetNext(posTarg);
			if(pAttrOri->m_idAttr==pAttrTarg->m_idAttr){
				if(pAttrTarg->m_nState==Update&&
					(pAttrTarg->m_nValType!=pAttrOri->m_nValType||
					pAttrTarg->m_nValType==CAPP_STRING&&pAttrTarg->m_nValSpecific!=pAttrOri->m_nValSpecific))
					AttrChangedList.AddTail(pAttrTarg);
				else
					break;
			}
		}
	}
	return CAPP_SUCCESS;
}

CAPP_RCODE CCAPPTreeView::PPB_UpdateClass_Attr_Col(CKBClass* pKBCls,CObList& AttrChangedList){
	if(!(PPBT_UpdateColumn(pKBCls,AttrChangedList)==CAPP_SUCCESS&&
		PPBT_UpdateClass_Attr(pKBCls)==CAPP_SUCCESS&&
		KBT_transact_commit()==CAPP_SUCCESS&&PBT_transact_commit()==CAPP_SUCCESS)){
		KBT_transact_rollback( );
		PBT_transact_rollback( );
		return CAPP_ERROR;
	}
	return CAPP_SUCCESS;
}
	
CAPP_RCODE CCAPPTreeView::PPBT_UpdateClass_Attr(CKBClass* pKBCls){
//首先更新所有属性列表
	POSITION position = pKBCls->m_cAttrList.GetHeadPosition();
	while(position)
	{
      CKBAttr *pcKBAttr = (CKBAttr *)pKBCls->m_cAttrList.GetNext(position);
      if(pcKBAttr->m_sName[0]=='\0')
         sprintf(pcKBAttr->m_sName,"attr%d",pcKBAttr->m_nRow);
      if(pcKBAttr->m_sAlias[0]=='\0')
         strcpy(pcKBAttr->m_sAlias,pcKBAttr->m_sName);
      if(pcKBAttr->m_nState==Update){
			CKBTAttr cKBTAttr;
			pcKBAttr->KB_Attr_to_TAttr(&cKBTAttr);
			if(cKBTAttr.KBT_updt_attr( )==CAPP_ERROR)
				return CAPP_ERROR;
		}
		if(pcKBAttr->m_nState==New){
         pcKBAttr->m_idCls=pKBCls->m_idCls;
			if(pcKBAttr->KB_add_attr() == CAPP_ERROR)				//此函数没有提交
			   return CAPP_ERROR;
      }
		if(pcKBAttr->m_nState==Del){
		   if(pcKBAttr->KB_del_attr() == CAPP_ERROR)					//此函数没有提交
			   return CAPP_ERROR;
		}
	}
//更新类
	CKBTClass cKBTCls;
	cKBTCls.m_idCls = pKBCls->m_idCls;
	cKBTCls.m_bCFF  = pKBCls->m_bCFF;
	if(cKBTCls.KBT_update_CFF()==CAPP_ERROR)
		return CAPP_ERROR;
/////////////////
	return CAPP_SUCCESS;
}

CAPP_RCODE CCAPPTreeView::PPBT_UpdateColumn(CKBClass* pKBCls,CObList& AttrChangedList){
//判断规程库中是否存在实例表
	BOOL bHaveFound=FALSE;			
	CObList PPBClsList;
   if(PPB_list_cls(&PPBClsList)==CAPP_ERROR){
		return CAPP_ERROR;
	}
	while(!PPBClsList.IsEmpty()){
		CPPBCls *pPPBCls=(CPPBCls *)PPBClsList.RemoveHead();
		if(pPPBCls->m_idCls==pKBCls->m_idCls)
			bHaveFound=TRUE;
		delete pPPBCls;
	}

//构造更新字段语句
	POSITION pos=AttrChangedList.GetHeadPosition();
	CKBAttr *pAttr;
	char sType[32];
	CAPP_RCODE rc;
	while(pos){
//修改知识库实例表
		pAttr= (CKBAttr *)pKBCls->m_cAttrList.GetNext(pos);
		CString strSql;
		if(g_nDataBaseType==SQL_SERVER||g_nDataBaseType==ACCESS)
			strSql.Format("ALTER TABLE KBT_INST_%I64d ALTER COLUMN INST_%I64d ",pKBCls->m_idCls,pAttr->m_idAttr);
		else if(g_nDataBaseType==ORACLE)
			strSql.Format("ALTER TABLE KBT_INST_%I64d MODIFY INST_%I64d ",pKBCls->m_idCls,pAttr->m_idAttr);
		else {
			strcpy(sCAPPErrMsg ,"不能兼容此数据库系统");
			return CAPP_ERROR;	
		}
		PPBT_GetAttrType(pAttr,sType);
		strSql+=sType;
		rc = SQLExecDirect(hstmt, (unsigned char*)strSql.GetBuffer(0), SQL_NTS);
		if(RETCODE_IS_FAILURE(rc)){
			SQLError(SQL_NULL_HENV,SQL_NULL_HDBC,hstmt,szSqlState,&dwNativeError,
               szErrorMsg,sizeof(szErrorMsg),&wErrorMsg);
			strcpy(sCAPPErrMsg ,(char *)(szErrorMsg));
			rc = ClearStmt(&hstmt);
			return CAPP_ERROR;	
		}
	   rc = ClearStmt(&hstmt);      
	   if(rc != SQL_SUCCESS)
			return CAPP_ERROR;
//修改规程库实例表
		if(bHaveFound){
			if(g_nDataBaseType==SQL_SERVER||g_nDataBaseType==ACCESS)
				strSql.Format("ALTER TABLE PPBT_%I64d ALTER COLUMN ATTR_%I64d ",pKBCls->m_idCls,pAttr->m_idAttr);
			else if(g_nDataBaseType==ORACLE)
				strSql.Format("ALTER TABLE PPBT_%I64d MODIFY ATTR_%I64d ",pKBCls->m_idCls,pAttr->m_idAttr);
			else {
				strcpy(sCAPPErrMsg ,"不能兼容此数据库系统");
				return CAPP_ERROR;	
			}
			PPBT_GetAttrType(pAttr,sType);
			strSql+=sType;
			rc = SQLExecDirect(ppb_hstmt, (unsigned char*)strSql.GetBuffer(0), SQL_NTS);
			if(RETCODE_IS_FAILURE(rc)){
				SQLError(SQL_NULL_HENV,SQL_NULL_HDBC,ppb_hstmt,szSqlState,&dwNativeError,
               szErrorMsg,sizeof(szErrorMsg),&wErrorMsg);
				strcpy(sCAPPErrMsg ,(char *)(szErrorMsg));
				rc = ClearPPBStmt(&ppb_hstmt);
				return CAPP_ERROR;	
			}
			rc = ClearPPBStmt(&ppb_hstmt);  
			if(rc != SQL_SUCCESS)
				return CAPP_ERROR;
		}
	}
	return CAPP_SUCCESS;
}
