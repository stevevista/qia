//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description:  Common utilities implementations
//-------------------------------------------------------------------------------------------------------

#include "utils.h"
#include "dispatch_object.h"

//-------------------------------------------------------------------------------------------------------

#define ERROR_MESSAGE_WIDE_MAXSIZE 1024
#define ERROR_MESSAGE_UTF8_MAXSIZE 2048

uint16_t *GetWin32ErroroMessage(uint16_t *buf, size_t buflen, Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t *bufptr = buf;
	size_t len;
	if (msg) {
		len = wcslen(msg);
		if (len >= buflen) len = buflen - 1;
		if (len > 0) memcpy(bufptr, msg, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 2) {
			bufptr[0] = ':';
			bufptr[1] = ' ';
			buflen -= 2;
			bufptr += 2;
		}
	}
	if (msg2) {
		len = wcslen(msg2);
		if (len >= buflen) len = buflen - 1;
		if (len > 0) memcpy(bufptr, msg2, len * sizeof(uint16_t));
		buflen -= len;
		bufptr += len;
		if (buflen > 1) {
			bufptr[0] = ' ';
			buflen -= 1;
			bufptr += 1;
		}
	}
	if (buflen > 1) {
		len = desc ? wcslen(desc) : 0;
		if (len > 0) {
			if (len >= buflen) len = buflen - 1;
			memcpy(bufptr, desc, len * sizeof(OLECHAR));
		}
		else {
			len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, hrcode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPOLESTR)bufptr, (DWORD)buflen - 1, 0);
			if (len == 0) len = swprintf_s((LPOLESTR)bufptr, buflen - 1, L"Error 0x%08X", hrcode);
		}
		buflen -= len;
		bufptr += len;
	}
	if (buflen > 0) bufptr[0] = 0;
	return buf;
}

char *GetWin32ErroroMessage(char *buf, size_t buflen, Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t buf_wide[ERROR_MESSAGE_WIDE_MAXSIZE];
	GetWin32ErroroMessage(buf_wide, ERROR_MESSAGE_WIDE_MAXSIZE, isolate, hrcode, msg, msg2, desc);
	int rcode = WideCharToMultiByte(CP_UTF8, 0, (WCHAR*)buf_wide, -1, buf, buflen, NULL, NULL);
	if (rcode < 0) rcode = 0;
	buf[rcode] = 0;
	return buf;
}

Local<String> GetWin32ErroroMessage(Isolate *isolate, HRESULT hrcode, LPCOLESTR msg, LPCOLESTR msg2, LPCOLESTR desc) {
	uint16_t buf_wide[ERROR_MESSAGE_WIDE_MAXSIZE];
	return String::NewFromTwoByte(isolate, GetWin32ErroroMessage(buf_wide, ERROR_MESSAGE_WIDE_MAXSIZE, isolate, hrcode, msg, msg2, desc));
}

//-------------------------------------------------------------------------------------------------------

Local<Value> Variant2Array(Isolate *isolate, const VARIANT &v) {
	if ((v.vt & VT_ARRAY) == 0) return Null(isolate);
	SAFEARRAY *varr = (v.vt & VT_BYREF) != 0 ? *v.pparray : v.parray;
	if (!varr || varr->cDims != 1) return Null(isolate);
	VARTYPE vt = v.vt & VT_TYPEMASK;
	LONG cnt = (LONG)varr->rgsabound[0].cElements;
	Local<Array> arr = Array::New(isolate, cnt);
	for (LONG i = varr->rgsabound[0].lLbound; i < cnt; i++) {
		CComVariant vi;
		if SUCCEEDED(SafeArrayGetElement(varr, &i, (vt == VT_VARIANT) ? (void*)&vi : (void*)&vi.byref)) {
			if (vt != VT_VARIANT) vi.vt = vt;
			arr->Set((uint32_t)i, Variant2Value(isolate, vi, true));
		}
	}
	return arr;
}

Local<Value> Variant2Value(Isolate *isolate, const VARIANT &v, bool allow_disp) {
	if ((v.vt & VT_ARRAY) != 0) return Variant2Array(isolate, v);
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_NULL:
		return Null(isolate);
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		return Int32::New(isolate, by_ref ? *v.plVal : v.lVal);
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		return Uint32::New(isolate, by_ref ? *v.pulVal : v.ulVal);
	case VT_CY:
		return Number::New(isolate, (double)(by_ref ? v.pcyVal : &v.cyVal)->int64 / 10000.);
	case VT_R4:
		return Number::New(isolate, by_ref ? *v.pfltVal : v.fltVal);
	case VT_R8:
		return Number::New(isolate, by_ref ? *v.pdblVal : v.dblVal);
	case VT_DATE:
                return Date::New(isolate, FromOleDate(by_ref ? *v.pdate : v.date));
	case VT_DECIMAL: {
		DOUBLE dblval;
		if FAILED(VarR8FromDec(by_ref ? v.pdecVal : &v.decVal, &dblval)) return Undefined(isolate);
		return Number::New(isolate, dblval);		
	}
	case VT_BOOL:
		return Boolean::New(isolate, (by_ref ? *v.pboolVal : v.boolVal) == VARIANT_TRUE);
	case VT_DISPATCH: {
		IDispatch *disp = (by_ref ? *v.ppdispVal : v.pdispVal);
		if (!disp) return Null(isolate);
		if (allow_disp) {
			DispObjectImpl *impl;
			if (disp->QueryInterface(CLSID_DispObjectImpl, (void**)&impl) == S_OK) {
				return impl->obj.Get(isolate);
			}
			return DispObject::NodeCreate(isolate, disp, L"Dispatch", option_auto);
		}
		return String::NewFromUtf8(isolate, "[Dispatch]");
	}
	case VT_UNKNOWN: {
		CComPtr<IDispatch> disp;
		if (allow_disp && UnknownDispGet(by_ref ? *v.ppunkVal : v.punkVal, &disp)) {
			return DispObject::NodeCreate(isolate, disp, L"Unknown", option_auto);
		}
		return String::NewFromUtf8(isolate, "[Unknown]");
	}
	case VT_BSTR: {
        BSTR bstr = by_ref ? (v.pbstrVal ? *v.pbstrVal : nullptr) : v.bstrVal;
        if (!bstr) return String::Empty(isolate);
		return String::NewFromTwoByte(isolate, (uint16_t*)bstr);
    }
	case VT_VARIANT: 
		if (v.pvarVal) return Variant2Value(isolate, *v.pvarVal, allow_disp);
	}
	return Undefined(isolate);
}

Local<Value> Variant2String(Isolate *isolate, const VARIANT &v) {
	char buf[256] = {};
	VARTYPE vt = (v.vt & VT_TYPEMASK);
	bool by_ref = (v.vt & VT_BYREF) != 0;
	switch (vt) {
	case VT_EMPTY:
		strcpy(buf, "EMPTY");
		break;
	case VT_NULL:
		strcpy(buf, "NULL");
		break;
	case VT_I1:
	case VT_I2:
	case VT_I4:
	case VT_INT:
		sprintf_s(buf, "%i", (int)(by_ref ? *v.plVal : v.lVal));
		break;
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_UINT:
		sprintf_s(buf, "%u", (unsigned int)(by_ref ? *v.pulVal : v.ulVal));
		break;
	case VT_CY:
		sprintf_s(buf, "%03f", (double)(by_ref ? v.pcyVal : &v.cyVal)->int64 / 10000.);
		break;
	case VT_R4:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pfltVal : v.fltVal));
		break;
	case VT_R8:
		sprintf_s(buf, "%f", (double)(by_ref ? *v.pdblVal : v.dblVal));
		break;
	case VT_DATE:
		return Date::New(isolate, FromOleDate(by_ref ? *v.pdate : v.date));
	case VT_DECIMAL: {
		DOUBLE dblval;
		if FAILED(VarR8FromDec(by_ref ? v.pdecVal : &v.decVal, &dblval)) return Undefined(isolate); 
		sprintf_s(buf, "%f", (double)dblval);
		break;		
	}
	case VT_BOOL:
		strcpy(buf, ((by_ref ? *v.pboolVal : v.boolVal) == VARIANT_TRUE) ? "true" : "false");
	case VT_DISPATCH:
		strcpy(buf, "[Dispatch]");
		break;
	case VT_UNKNOWN: 
		strcpy(buf, "[Unknown]");
		break;
	case VT_VARIANT:
		if (v.pvarVal) return Variant2String(isolate, *v.pvarVal);
		break;
	default:
		CComVariant tmp;
		if (SUCCEEDED(VariantChangeType(&tmp, &v, 0, VT_BSTR)) && tmp.vt == VT_BSTR && v.bstrVal != nullptr) {
			return String::NewFromTwoByte(isolate, (uint16_t*)v.bstrVal);
		}
	}
	return String::NewFromUtf8(isolate, buf, String::kNormalString);
}

void Value2Variant(Isolate *isolate, Local<Value> &val, VARIANT &var, VARTYPE vt) {
	if (val.IsEmpty() || val->IsUndefined()) {
		var.vt = VT_EMPTY;
	}
	else if (val->IsNull()) {
		var.vt = VT_NULL;
	}
	else if (val->IsInt32()) {
		var.vt = VT_I4;
		var.lVal = val->Int32Value();
	}
	else if (val->IsUint32()) {
		var.ulVal = val->Uint32Value();
		var.vt = (var.ulVal <= 0x7FFFFFFF) ? VT_I4 : VT_UI4;
	}
	else if (val->IsNumber()) {
		var.vt = VT_R8;
		var.dblVal = val->NumberValue();
	}
	else if (val->IsDate()) {
		var.vt = VT_DATE;
		var.date = ToOleDate(val->NumberValue());
	}
	else if (val->IsBoolean()) {
		var.vt = VT_BOOL;
		var.boolVal = val->BooleanValue() ? VARIANT_TRUE : VARIANT_FALSE;
	}
	else if (val->IsArray() && (vt != VT_NULL)) {
		Local<Array> arr = v8::Local<Array>::Cast(val);
		uint32_t len = arr->Length();
		if (vt == VT_EMPTY) vt = VT_VARIANT;
		var.vt = VT_ARRAY | vt;
		var.parray = SafeArrayCreateVector(vt, 0, len);
		for (uint32_t i = 0; i < len; i++) {
			CComVariant v;
			Value2Variant(isolate, arr->Get(i), v, vt);
			void *pv;
			if (vt == VT_VARIANT) pv = (void*)&v;
			else if (vt == VT_DISPATCH || vt == VT_UNKNOWN || vt == VT_BSTR) pv = v.byref;
			else pv = (void*)&v.byref;
			SafeArrayPutElement(var.parray, (LONG*)&i, pv);
		}
		vt = VT_EMPTY;
	}
	else if (val->IsObject()) {
		Local<Object> obj = val->ToObject();
		if (!DispObject::GetValueOf(isolate, obj, var) && !VariantObject::GetValueOf(isolate, obj, var)) {
			var.vt = VT_DISPATCH;
			var.pdispVal = new DispObjectImpl(obj);
			var.pdispVal->AddRef();
		}
	}
	else {
		String::Value str(val);
		var.vt = VT_BSTR;
		var.bstrVal = (str.length() > 0) ? SysAllocString((LPOLESTR)*str) : 0;
	}
	if (vt != VT_EMPTY && vt != VT_NULL && vt != VT_VARIANT) {
		if FAILED(VariantChangeType(&var, &var, 0, vt))
			VariantClear(&var);
	}
}

bool Value2Unknown(Isolate *isolate, Local<Value> &val, IUnknown **unk) {
    if (val.IsEmpty() || !val->IsObject()) return false;
    Local<Object> obj = val->ToObject();
    CComVariant var;
    if (!DispObject::GetValueOf(isolate, obj, var) && !VariantObject::GetValueOf(isolate, obj, var)) return false;
    return VariantUnkGet(&var, unk);
}

bool UnknownDispGet(IUnknown *unk, IDispatch **disp) {
	if (!unk) return false;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IDispatch), (void**)disp)) {
		return true;
	}
	CComPtr<IEnumVARIANT> enum_ptr;
	if SUCCEEDED(unk->QueryInterface(__uuidof(IEnumVARIANT), (void**)&enum_ptr)) {
		*disp = new DispEnumImpl(enum_ptr);
		(*disp)->AddRef();
		return true;
	}
	return false;
}

bool VariantUnkGet(VARIANT *v, IUnknown **punk) {
	IUnknown *unk = NULL;
    if ((v->vt & VT_TYPEMASK) == VT_DISPATCH) {
		unk = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
    }
    else if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN) {
        unk = ((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal;
    }
	if (!unk) return false;
	unk->AddRef();
	*punk = unk;
	return true;
}

bool VariantDispGet(VARIANT *v, IDispatch **pdisp) {
	/*
	if ((v->vt & VT_ARRAY) != 0) {
		*disp = new DispArrayImpl(*v);
		(*disp)->AddRef();
		return true;
	}
	*/
	if ((v->vt & VT_TYPEMASK) == VT_DISPATCH) {
		IDispatch *disp = ((v->vt & VT_BYREF) != 0) ? *v->ppdispVal : v->pdispVal;
		if (!disp) return false;
        disp->AddRef();
		*pdisp = disp;
        return true;
    }
    if ((v->vt & VT_TYPEMASK) == VT_UNKNOWN) {
		return UnknownDispGet(((v->vt & VT_BYREF) != 0) ? *v->ppunkVal : v->punkVal, pdisp);
    }
    return false;
}

/*
* Microsoft OLE Date type:
* https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2008/82ab7w69(v=vs.90)
*/

double FromOleDate(double oleDate) {
    double posixDate = oleDate - 25569; // days from 1899 dec 30
    posixDate *= 24 * 60 * 60 * 1000;   // days to milliseconds
    return posixDate;
}

double ToOleDate(double posixDate) {
    double oleDate = posixDate / (24 * 60 * 60 * 1000); // milliseconds to days
    oleDate += 25569;                                   // days from 1899 dec 30
    return oleDate;
}

//-------------------------------------------------------------------------------------------------------
