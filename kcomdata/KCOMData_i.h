/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Sep 30 15:58:42 1999
 */
/* Compiler settings for C:\wxz\kcomdata\KCOMData.odl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __KCOMData_i_h__
#define __KCOMData_i_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IGridColumns_FWD_DEFINED__
#define __IGridColumns_FWD_DEFINED__
typedef interface IGridColumns IGridColumns;
#endif 	/* __IGridColumns_FWD_DEFINED__ */


#ifndef __IGridColumn_FWD_DEFINED__
#define __IGridColumn_FWD_DEFINED__
typedef interface IGridColumn IGridColumn;
#endif 	/* __IGridColumn_FWD_DEFINED__ */


#ifndef __IDropDownColumns_FWD_DEFINED__
#define __IDropDownColumns_FWD_DEFINED__
typedef interface IDropDownColumns IDropDownColumns;
#endif 	/* __IDropDownColumns_FWD_DEFINED__ */


#ifndef __IDropDownColumn_FWD_DEFINED__
#define __IDropDownColumn_FWD_DEFINED__
typedef interface IDropDownColumn IDropDownColumn;
#endif 	/* __IDropDownColumn_FWD_DEFINED__ */


#ifndef __ISelBookmarks_FWD_DEFINED__
#define __ISelBookmarks_FWD_DEFINED__
typedef interface ISelBookmarks ISelBookmarks;
#endif 	/* __ISelBookmarks_FWD_DEFINED__ */


#ifndef ___DKCOMRichGrid_FWD_DEFINED__
#define ___DKCOMRichGrid_FWD_DEFINED__
typedef interface _DKCOMRichGrid _DKCOMRichGrid;
#endif 	/* ___DKCOMRichGrid_FWD_DEFINED__ */


#ifndef ___DKCOMRichGridEvents_FWD_DEFINED__
#define ___DKCOMRichGridEvents_FWD_DEFINED__
typedef interface _DKCOMRichGridEvents _DKCOMRichGridEvents;
#endif 	/* ___DKCOMRichGridEvents_FWD_DEFINED__ */


#ifndef __KCOMRichGrid_FWD_DEFINED__
#define __KCOMRichGrid_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMRichGrid KCOMRichGrid;
#else
typedef struct KCOMRichGrid KCOMRichGrid;
#endif /* __cplusplus */

#endif 	/* __KCOMRichGrid_FWD_DEFINED__ */


#ifndef ___DKCOMRichCombo_FWD_DEFINED__
#define ___DKCOMRichCombo_FWD_DEFINED__
typedef interface _DKCOMRichCombo _DKCOMRichCombo;
#endif 	/* ___DKCOMRichCombo_FWD_DEFINED__ */


#ifndef ___DKCOMRichComboEvents_FWD_DEFINED__
#define ___DKCOMRichComboEvents_FWD_DEFINED__
typedef interface _DKCOMRichComboEvents _DKCOMRichComboEvents;
#endif 	/* ___DKCOMRichComboEvents_FWD_DEFINED__ */


#ifndef __KCOMRichCombo_FWD_DEFINED__
#define __KCOMRichCombo_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMRichCombo KCOMRichCombo;
#else
typedef struct KCOMRichCombo KCOMRichCombo;
#endif /* __cplusplus */

#endif 	/* __KCOMRichCombo_FWD_DEFINED__ */


#ifndef ___DKCOMRichDropDown_FWD_DEFINED__
#define ___DKCOMRichDropDown_FWD_DEFINED__
typedef interface _DKCOMRichDropDown _DKCOMRichDropDown;
#endif 	/* ___DKCOMRichDropDown_FWD_DEFINED__ */


#ifndef ___DKCOMRichDropDownEvents_FWD_DEFINED__
#define ___DKCOMRichDropDownEvents_FWD_DEFINED__
typedef interface _DKCOMRichDropDownEvents _DKCOMRichDropDownEvents;
#endif 	/* ___DKCOMRichDropDownEvents_FWD_DEFINED__ */


#ifndef __KCOMRichDropDown_FWD_DEFINED__
#define __KCOMRichDropDown_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMRichDropDown KCOMRichDropDown;
#else
typedef struct KCOMRichDropDown KCOMRichDropDown;
#endif /* __cplusplus */

#endif 	/* __KCOMRichDropDown_FWD_DEFINED__ */


#ifndef __KCOMGridColumns_FWD_DEFINED__
#define __KCOMGridColumns_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMGridColumns KCOMGridColumns;
#else
typedef struct KCOMGridColumns KCOMGridColumns;
#endif /* __cplusplus */

#endif 	/* __KCOMGridColumns_FWD_DEFINED__ */


#ifndef __KCOMGridColumn_FWD_DEFINED__
#define __KCOMGridColumn_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMGridColumn KCOMGridColumn;
#else
typedef struct KCOMGridColumn KCOMGridColumn;
#endif /* __cplusplus */

#endif 	/* __KCOMGridColumn_FWD_DEFINED__ */


#ifndef __DropDownColumns_FWD_DEFINED__
#define __DropDownColumns_FWD_DEFINED__

#ifdef __cplusplus
typedef class DropDownColumns DropDownColumns;
#else
typedef struct DropDownColumns DropDownColumns;
#endif /* __cplusplus */

#endif 	/* __DropDownColumns_FWD_DEFINED__ */


#ifndef __DropDownColumn_FWD_DEFINED__
#define __DropDownColumn_FWD_DEFINED__

#ifdef __cplusplus
typedef class DropDownColumn DropDownColumn;
#else
typedef struct DropDownColumn DropDownColumn;
#endif /* __cplusplus */

#endif 	/* __DropDownColumn_FWD_DEFINED__ */


#ifndef __SelBookmarks_FWD_DEFINED__
#define __SelBookmarks_FWD_DEFINED__

#ifdef __cplusplus
typedef class SelBookmarks SelBookmarks;
#else
typedef struct SelBookmarks SelBookmarks;
#endif /* __cplusplus */

#endif 	/* __SelBookmarks_FWD_DEFINED__ */


void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __KCOMDATALib_LIBRARY_DEFINED__
#define __KCOMDATALib_LIBRARY_DEFINED__

/* library KCOMDATALib */
/* [control][helpstring][helpfile][version][uuid] */ 






typedef 
enum _DataMode
    {	DataModeBound	= 0,
	DataModeAddItem	= 1
    }	Constants_DataMode;

typedef 
enum _Style
    {	StyleEdit	= 0,
	StyleEditButton	= 1,
	StyleCheckBox	= 2,
	StyleComboBox	= 3,
	StyleButton	= 4
    }	Constants_Style;

typedef 
enum _Case
    {	CaseUnchanged	= 0,
	CaseUpper	= 1,
	CaseLower	= 2
    }	Constants_Case;

typedef 
enum _DataType
    {	DataTypeText	= 0,
	DataTypeVARIANT_BOOL	= 1,
	DataTypeByte	= 2,
	DataTypeInteger	= 3,
	DataTypeLong	= 4,
	DataTypeSingle	= 5,
	DataTypeCurrency	= 6,
	DataTypeDate	= 7
    }	Constants_DataType;

typedef 
enum _CaptionAlignment
    {	CaptionAlignmentLeft	= 0,
	CaptionAlignmentRight	= 1,
	CaptionAlignmentCenter	= 2
    }	Constants_CaptionAlignment;

typedef 
enum _DividerType
    {	DividerTypeNone	= 0,
	DividerTypeVertical	= 1,
	DividerTypeHorizontal	= 2,
	DividerTypeBoth	= 3
    }	Constants_DividerType;

typedef 
enum _DividerStyle
    {	DividerStyleSolidline	= 0,
	DividerStyleDashline	= 1,
	DividerStyleDot	= 2,
	DividerStyleDashDot	= 3,
	DividerStyleDashDotDot	= 4
    }	Constants_DividerStyle;

typedef 
enum _ColumnCaptionAlignment
    {	ColCapAlignLeftJustify	= 0,
	ColCapAlignRightJustify	= 1,
	ColCapAlignCenter	= 2,
	ColCapAlignUseColumnAlignment	= 3
    }	Constants_ColumnCaptionAlignment;

typedef 
enum _ComboStyle
    {	ComboStyleDropDown	= 0,
	ComboStyleDropList	= 1
    }	Constants_ComboStyle;


EXTERN_C const IID LIBID_KCOMDATALib;

#ifndef __IGridColumns_DISPINTERFACE_DEFINED__
#define __IGridColumns_DISPINTERFACE_DEFINED__

/* dispinterface IGridColumns */
/* [uuid] */ 


EXTERN_C const IID DIID_IGridColumns;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("61724100-12B3-11D3-A7FE-0080C8763FA4")
    IGridColumns : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IGridColumnsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IGridColumns __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IGridColumns __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IGridColumns __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IGridColumns __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IGridColumns __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IGridColumns __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IGridColumns __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IGridColumnsVtbl;

    interface IGridColumns
    {
        CONST_VTBL struct IGridColumnsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGridColumns_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IGridColumns_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IGridColumns_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IGridColumns_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IGridColumns_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IGridColumns_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IGridColumns_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IGridColumns_DISPINTERFACE_DEFINED__ */


#ifndef __IGridColumn_DISPINTERFACE_DEFINED__
#define __IGridColumn_DISPINTERFACE_DEFINED__

/* dispinterface IGridColumn */
/* [uuid] */ 


EXTERN_C const IID DIID_IGridColumn;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("61724103-12B3-11D3-A7FE-0080C8763FA4")
    IGridColumn : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IGridColumnVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IGridColumn __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IGridColumn __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IGridColumn __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IGridColumn __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IGridColumn __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IGridColumn __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IGridColumn __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IGridColumnVtbl;

    interface IGridColumn
    {
        CONST_VTBL struct IGridColumnVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGridColumn_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IGridColumn_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IGridColumn_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IGridColumn_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IGridColumn_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IGridColumn_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IGridColumn_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IGridColumn_DISPINTERFACE_DEFINED__ */


#ifndef __IDropDownColumns_DISPINTERFACE_DEFINED__
#define __IDropDownColumns_DISPINTERFACE_DEFINED__

/* dispinterface IDropDownColumns */
/* [uuid] */ 


EXTERN_C const IID DIID_IDropDownColumns;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("491325C0-3068-11D3-A7FE-0080C8763FA4")
    IDropDownColumns : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IDropDownColumnsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IDropDownColumns __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IDropDownColumns __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IDropDownColumns __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IDropDownColumns __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IDropDownColumns __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IDropDownColumns __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IDropDownColumns __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IDropDownColumnsVtbl;

    interface IDropDownColumns
    {
        CONST_VTBL struct IDropDownColumnsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDropDownColumns_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDropDownColumns_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDropDownColumns_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDropDownColumns_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IDropDownColumns_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IDropDownColumns_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IDropDownColumns_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IDropDownColumns_DISPINTERFACE_DEFINED__ */


#ifndef __IDropDownColumn_DISPINTERFACE_DEFINED__
#define __IDropDownColumn_DISPINTERFACE_DEFINED__

/* dispinterface IDropDownColumn */
/* [uuid] */ 


EXTERN_C const IID DIID_IDropDownColumn;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("491325C3-3068-11D3-A7FE-0080C8763FA4")
    IDropDownColumn : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IDropDownColumnVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IDropDownColumn __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IDropDownColumn __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IDropDownColumn __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IDropDownColumn __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IDropDownColumn __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IDropDownColumn __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IDropDownColumn __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IDropDownColumnVtbl;

    interface IDropDownColumn
    {
        CONST_VTBL struct IDropDownColumnVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDropDownColumn_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDropDownColumn_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDropDownColumn_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDropDownColumn_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IDropDownColumn_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IDropDownColumn_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IDropDownColumn_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IDropDownColumn_DISPINTERFACE_DEFINED__ */


#ifndef __ISelBookmarks_DISPINTERFACE_DEFINED__
#define __ISelBookmarks_DISPINTERFACE_DEFINED__

/* dispinterface ISelBookmarks */
/* [uuid] */ 


EXTERN_C const IID DIID_ISelBookmarks;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("8119E220-3B61-11D3-A7FE-0080C8763FA4")
    ISelBookmarks : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ISelBookmarksVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ISelBookmarks __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ISelBookmarks __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ISelBookmarks __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ISelBookmarks __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ISelBookmarks __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ISelBookmarks __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ISelBookmarks __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } ISelBookmarksVtbl;

    interface ISelBookmarks
    {
        CONST_VTBL struct ISelBookmarksVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISelBookmarks_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISelBookmarks_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISelBookmarks_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISelBookmarks_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ISelBookmarks_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ISelBookmarks_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ISelBookmarks_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ISelBookmarks_DISPINTERFACE_DEFINED__ */


#ifndef ___DKCOMRichGrid_DISPINTERFACE_DEFINED__
#define ___DKCOMRichGrid_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichGrid */
/* [hidden][helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichGrid;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("0CFC2323-12A9-11D3-A7FE-0080C8763FA4")
    _DKCOMRichGrid : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichGridVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichGrid __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichGrid __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichGrid __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichGrid __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichGrid __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichGrid __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichGrid __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichGridVtbl;

    interface _DKCOMRichGrid
    {
        CONST_VTBL struct _DKCOMRichGridVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichGrid_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichGrid_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichGrid_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichGrid_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichGrid_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichGrid_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichGrid_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichGrid_DISPINTERFACE_DEFINED__ */


#ifndef ___DKCOMRichGridEvents_DISPINTERFACE_DEFINED__
#define ___DKCOMRichGridEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichGridEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichGridEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("0CFC2324-12A9-11D3-A7FE-0080C8763FA4")
    _DKCOMRichGridEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichGridEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichGridEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichGridEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichGridEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichGridEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichGridEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichGridEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichGridEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichGridEventsVtbl;

    interface _DKCOMRichGridEvents
    {
        CONST_VTBL struct _DKCOMRichGridEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichGridEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichGridEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichGridEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichGridEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichGridEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichGridEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichGridEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichGridEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_KCOMRichGrid;

#ifdef __cplusplus

class DECLSPEC_UUID("0CFC2325-12A9-11D3-A7FE-0080C8763FA4")
KCOMRichGrid;
#endif

#ifndef ___DKCOMRichCombo_DISPINTERFACE_DEFINED__
#define ___DKCOMRichCombo_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichCombo */
/* [hidden][helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichCombo;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("029288C7-2FFF-11D3-B446-0080C8F18522")
    _DKCOMRichCombo : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichComboVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichCombo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichCombo __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichCombo __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichCombo __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichCombo __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichCombo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichCombo __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichComboVtbl;

    interface _DKCOMRichCombo
    {
        CONST_VTBL struct _DKCOMRichComboVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichCombo_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichCombo_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichCombo_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichCombo_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichCombo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichCombo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichCombo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichCombo_DISPINTERFACE_DEFINED__ */


#ifndef ___DKCOMRichComboEvents_DISPINTERFACE_DEFINED__
#define ___DKCOMRichComboEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichComboEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichComboEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("029288C8-2FFF-11D3-B446-0080C8F18522")
    _DKCOMRichComboEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichComboEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichComboEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichComboEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichComboEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichComboEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichComboEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichComboEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichComboEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichComboEventsVtbl;

    interface _DKCOMRichComboEvents
    {
        CONST_VTBL struct _DKCOMRichComboEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichComboEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichComboEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichComboEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichComboEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichComboEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichComboEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichComboEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichComboEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_KCOMRichCombo;

#ifdef __cplusplus

class DECLSPEC_UUID("029288C9-2FFF-11D3-B446-0080C8F18522")
KCOMRichCombo;
#endif

#ifndef ___DKCOMRichDropDown_DISPINTERFACE_DEFINED__
#define ___DKCOMRichDropDown_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichDropDown */
/* [hidden][helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichDropDown;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("029288CB-2FFF-11D3-B446-0080C8F18522")
    _DKCOMRichDropDown : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichDropDownVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichDropDown __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichDropDown __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichDropDown __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichDropDown __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichDropDown __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichDropDown __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichDropDown __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichDropDownVtbl;

    interface _DKCOMRichDropDown
    {
        CONST_VTBL struct _DKCOMRichDropDownVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichDropDown_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichDropDown_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichDropDown_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichDropDown_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichDropDown_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichDropDown_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichDropDown_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichDropDown_DISPINTERFACE_DEFINED__ */


#ifndef ___DKCOMRichDropDownEvents_DISPINTERFACE_DEFINED__
#define ___DKCOMRichDropDownEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMRichDropDownEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMRichDropDownEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("029288CC-2FFF-11D3-B446-0080C8F18522")
    _DKCOMRichDropDownEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMRichDropDownEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMRichDropDownEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMRichDropDownEventsVtbl;

    interface _DKCOMRichDropDownEvents
    {
        CONST_VTBL struct _DKCOMRichDropDownEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMRichDropDownEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMRichDropDownEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMRichDropDownEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMRichDropDownEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMRichDropDownEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMRichDropDownEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMRichDropDownEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMRichDropDownEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_KCOMRichDropDown;

#ifdef __cplusplus

class DECLSPEC_UUID("029288CD-2FFF-11D3-B446-0080C8F18522")
KCOMRichDropDown;
#endif

EXTERN_C const CLSID CLSID_KCOMGridColumns;

#ifdef __cplusplus

class DECLSPEC_UUID("61724102-12B3-11D3-A7FE-0080C8763FA4")
KCOMGridColumns;
#endif

EXTERN_C const CLSID CLSID_KCOMGridColumn;

#ifdef __cplusplus

class DECLSPEC_UUID("61724105-12B3-11D3-A7FE-0080C8763FA4")
KCOMGridColumn;
#endif

EXTERN_C const CLSID CLSID_DropDownColumns;

#ifdef __cplusplus

class DECLSPEC_UUID("491325C2-3068-11D3-A7FE-0080C8763FA4")
DropDownColumns;
#endif

EXTERN_C const CLSID CLSID_DropDownColumn;

#ifdef __cplusplus

class DECLSPEC_UUID("491325C5-3068-11D3-A7FE-0080C8763FA4")
DropDownColumn;
#endif

EXTERN_C const CLSID CLSID_SelBookmarks;

#ifdef __cplusplus

class DECLSPEC_UUID("8119E222-3B61-11D3-A7FE-0080C8763FA4")
SelBookmarks;
#endif
#endif /* __KCOMDATALib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
