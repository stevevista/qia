//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: Common utilities for translation COM - NodeJS
//-------------------------------------------------------------------------------------------------------

#pragma once

#include "unknown_objects.h"

//-------------------------------------------------------------------------------------------------------

#ifdef _DEBUG
#define NODE_DEBUG
#endif

#ifdef NODE_DEBUG
#define NODE_DEBUG_PREFIX "### "
#define NODE_DEBUG_MSG(msg) { printf(NODE_DEBUG_PREFIX"%s", msg); std::cout << std::endl; }
#define NODE_DEBUG_FMT(msg, arg) { std::cout << NODE_DEBUG_PREFIX; printf(msg, arg); std::cout << std::endl; }
#define NODE_DEBUG_FMT2(msg, arg, arg2) { std::cout << NODE_DEBUG_PREFIX; printf(msg, arg, arg2); std::cout << std::endl; }
#else
#define NODE_DEBUG_MSG(msg)
#define NODE_DEBUG_FMT(msg, arg)
#define NODE_DEBUG_FMT2(msg, arg, arg2)
#endif

//-------------------------------------------------------------------------------------------------------
#ifndef USE_ATL


#endif
//-------------------------------------------------------------------------------------------------------

Local<String> GetWin32ErroroMessage(Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2 = 0, LPCOLESTR desc = 0);

inline Local<Value> Win32Error(Isolate *isolate, HRESULT hrcode, LPCOLESTR id = 0, LPCOLESTR msg = 0) {
	Local<Value> err = Exception::Error(GetWin32ErroroMessage(isolate, hrcode, id, msg));
	Local<Object> obj = err->ToObject();
	obj->Set(String::NewFromUtf8(isolate, "errno"), Integer::New(isolate, hrcode));
	return err;
}

inline Local<Value> DispError(Isolate *isolate, HRESULT hrcode, LPCOLESTR id = 0, LPCOLESTR msg = 0, EXCEPINFO *except = 0) {
    CComBSTR desc;
    CComPtr<IErrorInfo> errinfo;
    HRESULT hr = GetErrorInfo(0, &errinfo);
    if (hr == S_OK) errinfo->GetDescription(&desc);
	Local<Value> err = Exception::Error(GetWin32ErroroMessage(isolate, hrcode, id, msg, desc));
	Local<Object> obj = err->ToObject();
	obj->Set(String::NewFromUtf8(isolate, "errno"), Integer::New(isolate, hrcode));
	if (except) {
		if (except->scode != 0) obj->Set(String::NewFromUtf8(isolate, "code"), Integer::New(isolate, except->scode));
		else if (except->wCode != 0) obj->Set(String::NewFromUtf8(isolate, "code"), Integer::New(isolate, except->wCode));
		if (except->bstrSource != 0) obj->Set(String::NewFromUtf8(isolate, "source"), String::NewFromTwoByte(isolate, (uint16_t*)except->bstrSource));
		if (except->bstrDescription != 0) obj->Set(String::NewFromUtf8(isolate, "description"), String::NewFromTwoByte(isolate, (uint16_t*)except->bstrDescription));
	}
	return err;
}

inline Local<Value> DispErrorNull(Isolate *isolate) {
    return Exception::TypeError(String::NewFromUtf8(isolate, "DispNull"));
}

inline Local<Value> DispErrorInvalid(Isolate *isolate) {
    return Exception::TypeError(String::NewFromUtf8(isolate, "DispInvalid"));
}

inline Local<Value> TypeError(Isolate *isolate, const char *msg) {
    return Exception::TypeError(String::NewFromUtf8(isolate, msg));
}

inline Local<Value> InvalidArgumentsError(Isolate *isolate) {
    return Exception::TypeError(String::NewFromUtf8(isolate, "Invalid arguments"));
}

inline Local<Value> Error(Isolate *isolate, const char *msg) {
    return Exception::Error(String::NewFromUtf8(isolate, msg));
}

//-------------------------------------------------------------------------------------------------------

inline HRESULT DispFind(IDispatch *disp, LPOLESTR name, DISPID *dispid) {
	LPOLESTR names[] = { name };
	return disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispid);
}

inline HRESULT DispInvoke(IDispatch *disp, DISPID dispid, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, EXCEPINFO *except = 0) {
	DISPPARAMS params = { args, 0, argcnt, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	if (flags == DISPATCH_PROPERTYPUT) { // It`s a magic
		params.cNamedArgs = 1;
		params.rgdispidNamedArgs = &dispidNamed;
	}
	return disp->Invoke(dispid, IID_NULL, 0, flags, &params, ret, except, 0);
}

inline HRESULT DispInvoke(IDispatch *disp, LPOLESTR name, UINT argcnt = 0, VARIANT *args = 0, VARIANT *ret = 0, WORD  flags = DISPATCH_METHOD, DISPID *dispid = 0, EXCEPINFO *except = 0) {
	LPOLESTR names[] = { name };
    DISPID dispids[] = { 0 };
	HRESULT hrcode = disp->GetIDsOfNames(GUID_NULL, names, 1, 0, dispids);
	if SUCCEEDED(hrcode) hrcode = DispInvoke(disp, dispids[0], argcnt, args, ret, flags, except);
	if (dispid) *dispid = dispids[0];
	return hrcode;
}

//-------------------------------------------------------------------------------------------------------

template<typename INTTYPE>
inline INTTYPE Variant2Int(const VARIANT &v, const INTTYPE def) {
    VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
    switch (vt) {
    case VT_EMPTY:
    case VT_NULL:
        return def;
    case VT_I1:
    case VT_I2:
    case VT_I4:
    case VT_INT:
        return (INTTYPE)(by_ref ? *v.plVal : v.lVal);
    case VT_UI1:
    case VT_UI2:
    case VT_UI4:
    case VT_UINT:
        return (INTTYPE)(by_ref ? *v.pulVal : v.ulVal);
	case VT_CY:
		return (INTTYPE)((by_ref ? v.pcyVal : &v.cyVal)->int64 / 10000);
	case VT_R4:
        return (INTTYPE)(by_ref ? *v.pfltVal : v.fltVal);
    case VT_R8:
        return (INTTYPE)(by_ref ? *v.pdblVal : v.dblVal);
    case VT_DATE:
        return (INTTYPE)(by_ref ? *v.pdate : v.date);
	case VT_DECIMAL: {
		LONG64 int64val;
		return SUCCEEDED(VarI8FromDec(by_ref ? v.pdecVal : &v.decVal, &int64val)) ? (INTTYPE)int64val : def; 
	}
	case VT_BOOL:
        return (v.boolVal == VARIANT_TRUE) ? 1 : 0;
	case VT_VARIANT:
		if (v.pvarVal) return Variant2Int<INTTYPE>(*v.pvarVal, def);
	}
    VARIANT dst;
    return SUCCEEDED(VariantChangeType(&dst, &v, 0, VT_INT)) ? (INTTYPE)dst.intVal : def;
}

Local<Value> Variant2Array(Isolate *isolate, const VARIANT &v);
Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v, bool allow_disp = false);
Local<Value> Variant2String(Isolate *isolate, const VARIANT &v);
void Value2Variant(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt = VT_EMPTY);
bool Value2Unknown(Isolate *isolate, Local<Value> &val, IUnknown **unk);
bool VariantUnkGet(VARIANT *v, IUnknown **unk);
bool VariantDispGet(VARIANT *v, IDispatch **disp);
bool UnknownDispGet(IUnknown *unk, IDispatch **disp);

//-------------------------------------------------------------------------------------------------------

inline bool v8val2bool(const Local<Value> &v, bool def) {
    if (v.IsEmpty()) return def;
    if (v->IsBoolean()) return v->BooleanValue();
    if (v->IsInt32()) return v->Int32Value() != 0;
    if (v->IsUint32()) return v->Uint32Value() != 0;
    return def;
}

//-------------------------------------------------------------------------------------------------------

class VarArguments {
public:
	std::vector<CComVariant> items;
	VarArguments() {}
	VarArguments(Isolate *isolate, Local<Value> value) {
		items.resize(1);
		Value2Variant(isolate, value, items[0]);
	}
	VarArguments(Isolate *isolate, const FunctionCallbackInfo<Value> &args) {
		int argcnt = args.Length();
		items.resize(argcnt);
		for (int i = 0; i < argcnt; i ++)
			Value2Variant(isolate, args[argcnt - i - 1], items[i]);
	}
};

class NodeArguments {
public:
	std::vector<Local<Value>> items;
	NodeArguments(Isolate *isolate, DISPPARAMS *pDispParams, bool allow_disp) {
		UINT argcnt = pDispParams->cArgs;
		items.resize(argcnt);
		for (UINT i = 0; i < argcnt; i++) {
			items[i] = Variant2Value(isolate, pDispParams->rgvarg[argcnt - i - 1], allow_disp);
		}
	}
};


double FromOleDate(double);
double ToOleDate(double);

//-------------------------------------------------------------------------------------------------------
