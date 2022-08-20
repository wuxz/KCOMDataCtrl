/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Fri May 07 23:33:08 1999
 */
/* Compiler settings for D:\wxz\KCOMDBGrid\KCOMDBGrid.odl:
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

#ifndef __KCOMDBGrid_i_h__
#define __KCOMDBGrid_i_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef ___DKCOMDBGrid_FWD_DEFINED__
#define ___DKCOMDBGrid_FWD_DEFINED__
typedef interface _DKCOMDBGrid _DKCOMDBGrid;
#endif 	/* ___DKCOMDBGrid_FWD_DEFINED__ */


#ifndef ___DKCOMDBGridEvents_FWD_DEFINED__
#define ___DKCOMDBGridEvents_FWD_DEFINED__
typedef interface _DKCOMDBGridEvents _DKCOMDBGridEvents;
#endif 	/* ___DKCOMDBGridEvents_FWD_DEFINED__ */


#ifndef __KCOMDBGrid_FWD_DEFINED__
#define __KCOMDBGrid_FWD_DEFINED__

#ifdef __cplusplus
typedef class KCOMDBGrid KCOMDBGrid;
#else
typedef struct KCOMDBGrid KCOMDBGrid;
#endif /* __cplusplus */

#endif 	/* __KCOMDBGrid_FWD_DEFINED__ */


void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __KCOMDBGRIDLib_LIBRARY_DEFINED__
#define __KCOMDBGRIDLib_LIBRARY_DEFINED__

/* library KCOMDBGRIDLib */
/* [control][helpstring][helpfile][version][uuid] */ 

typedef /* [public][public] */ 
enum __MIDL___MIDL_itf_KCOMDBGrid_0000_0001
    {	dmBind	= 0,
	dmManual	= 1
    }	DataBindMode;

typedef /* [public][public] */ 
enum __MIDL___MIDL_itf_KCOMDBGrid_0000_0002
    {	glsBoth	= 3,
	glsVert	= 2,
	glsHoriz	= 1,
	glsNone	= 0
    }	GLStyle;


EXTERN_C const IID LIBID_KCOMDBGRIDLib;

#ifndef ___DKCOMDBGrid_DISPINTERFACE_DEFINED__
#define ___DKCOMDBGrid_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMDBGrid */
/* [hidden][helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMDBGrid;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("AC212644-FBE2-11D2-A7FE-0080C8763FA4")
    _DKCOMDBGrid : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMDBGridVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMDBGrid __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMDBGrid __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMDBGrid __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMDBGrid __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMDBGrid __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMDBGrid __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMDBGrid __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMDBGridVtbl;

    interface _DKCOMDBGrid
    {
        CONST_VTBL struct _DKCOMDBGridVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMDBGrid_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMDBGrid_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMDBGrid_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMDBGrid_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMDBGrid_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMDBGrid_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMDBGrid_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMDBGrid_DISPINTERFACE_DEFINED__ */


#ifndef ___DKCOMDBGridEvents_DISPINTERFACE_DEFINED__
#define ___DKCOMDBGridEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DKCOMDBGridEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DKCOMDBGridEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("AC212645-FBE2-11D2-A7FE-0080C8763FA4")
    _DKCOMDBGridEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DKCOMDBGridEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _DKCOMDBGridEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _DKCOMDBGridEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _DKCOMDBGridEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _DKCOMDBGridEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _DKCOMDBGridEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _DKCOMDBGridEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _DKCOMDBGridEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _DKCOMDBGridEventsVtbl;

    interface _DKCOMDBGridEvents
    {
        CONST_VTBL struct _DKCOMDBGridEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DKCOMDBGridEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DKCOMDBGridEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DKCOMDBGridEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DKCOMDBGridEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DKCOMDBGridEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DKCOMDBGridEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DKCOMDBGridEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DKCOMDBGridEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_KCOMDBGrid;

#ifdef __cplusplus

class DECLSPEC_UUID("AC212646-FBE2-11D2-A7FE-0080C8763FA4")
KCOMDBGrid;
#endif
#endif /* __KCOMDBGRIDLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
