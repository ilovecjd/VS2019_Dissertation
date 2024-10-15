

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 12:14:07 2038
 */
/* Compiler settings for Simulator.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __Simulator_h_h__
#define __Simulator_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ISimulator_FWD_DEFINED__
#define __ISimulator_FWD_DEFINED__
typedef interface ISimulator ISimulator;

#endif 	/* __ISimulator_FWD_DEFINED__ */


#ifndef __Simulator_FWD_DEFINED__
#define __Simulator_FWD_DEFINED__

#ifdef __cplusplus
typedef class Simulator Simulator;
#else
typedef struct Simulator Simulator;
#endif /* __cplusplus */

#endif 	/* __Simulator_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __Simulator_LIBRARY_DEFINED__
#define __Simulator_LIBRARY_DEFINED__

/* library Simulator */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_Simulator;

#ifndef __ISimulator_DISPINTERFACE_DEFINED__
#define __ISimulator_DISPINTERFACE_DEFINED__

/* dispinterface ISimulator */
/* [uuid] */ 


EXTERN_C const IID DIID_ISimulator;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("7a36e8b6-03f2-4059-9cc1-3e7ff6a912c3")
    ISimulator : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ISimulatorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISimulator * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISimulator * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISimulator * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISimulator * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISimulator * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISimulator * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISimulator * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } ISimulatorVtbl;

    interface ISimulator
    {
        CONST_VTBL struct ISimulatorVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISimulator_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISimulator_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISimulator_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISimulator_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISimulator_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISimulator_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISimulator_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ISimulator_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Simulator;

#ifdef __cplusplus

class DECLSPEC_UUID("db374100-169a-4389-b9b4-c6ab5e02548b")
Simulator;
#endif
#endif /* __Simulator_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


