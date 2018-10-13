#pragma once

#include <SDKDDKVer.h>

// Windows Header Files:
#define WIN32_LEAN_AND_MEAN                     // Exclude rarely-used stuff from Windows headers
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit
#include <windows.h>

#include <ole2.h>
#include <ocidl.h>

// STD headers
#include <iostream>
#include <stdio.h>
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <initializer_list>

// Node JS headers
#include <v8.h>
#include <node.h>
#include <node_version.h>
#include <node_object_wrap.h>
#include <node_buffer.h>
using namespace v8;
using namespace node;

#include <nan.h>
#include <uv.h>


class CComVariant : public VARIANT {
public:
    inline CComVariant() { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
	}
    inline CComVariant(const CComVariant &src) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		VariantCopyInd(this, &src);
	}
    inline CComVariant(const VARIANT &src) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		VariantCopyInd(this, &src);
	}
    inline CComVariant(LONG v) { 
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		vt = VT_I4;
		lVal = v; 
	}
	inline CComVariant(LPOLESTR v) {
		memset((VARIANT*)this, 0, sizeof(VARIANT));
		vt = VT_BSTR;
		bstrVal = SysAllocString(v);
	}
	inline ~CComVariant() {
		Clear(); 
	}
    inline void Clear() {
        if (vt != VT_EMPTY)
            VariantClear(this);
    }
    inline void Detach(VARIANT *dst) {
        *dst = *this;
        vt = VT_EMPTY;
    }
	inline HRESULT CopyTo(VARIANT *dst) {
		return VariantCopy(dst, this);
	}

	inline HRESULT ChangeType(VARTYPE vtNew, const VARIANT* pSrc = NULL) {
		return VariantChangeType(this, pSrc ? pSrc : this, 0, vtNew);
	}

	inline ULONG ArrayLength() {
		if ((vt & VT_ARRAY) == 0) return 0;
		SAFEARRAY *varr = (vt & VT_BYREF) != 0 ? *pparray : parray;
		return varr ? varr->rgsabound[0].cElements : 0;
	}

	inline HRESULT ArrayGet(LONG index, CComVariant &var) {
		if ((vt & VT_ARRAY) == 0) return E_NOTIMPL;
		SAFEARRAY *varr = (vt & VT_BYREF) != 0 ? *pparray : parray;
		if (!varr) return E_FAIL;
		index += varr->rgsabound[0].lLbound;
		VARTYPE vart = vt & VT_TYPEMASK;
		HRESULT hr = SafeArrayGetElement(varr, &index, (vart == VT_VARIANT) ? (void*)&var : (void*)&var.byref);
		if (SUCCEEDED(hr) && vart != VT_VARIANT) var.vt = vart;
		return hr;
	}
	template<typename T>
	inline T* ArrayGet(ULONG index = 0) {
		return ((T*)parray->pvData) + index;
	}
	inline HRESULT ArrayCreate(VARTYPE avt, ULONG cnt) {
		Clear();
		parray = SafeArrayCreateVector(avt, 0, cnt);
		if (!parray) return E_UNEXPECTED;
		vt = VT_ARRAY | avt;
		return S_OK;
	}
	inline HRESULT ArrayResize(ULONG cnt) {
		SAFEARRAYBOUND bnds = { cnt, 0 };
		return SafeArrayRedim(parray, &bnds);
	}
};

class CComBSTR {
public:
    BSTR p;
    inline CComBSTR() : p(0) {}
    inline CComBSTR(const CComBSTR &src) : p(0) {}
    inline ~CComBSTR() { Free(); }
    inline void Attach(BSTR _p) { Free(); p = _p; }
    inline BSTR Detach() { BSTR pp = p; p = 0; return pp; }
    inline void Free() { if (p) { SysFreeString(p); p = 0; } }

    inline operator BSTR () const { return p; }
    inline BSTR* operator&() { return &p; }
    inline bool operator!() const { return (p == 0); }
    inline bool operator!=(BSTR _p) const { return !operator==(_p); }
    inline bool operator==(BSTR _p) const { return p == _p; }
    inline BSTR operator = (BSTR _p) {
        if (p != _p) Attach(_p ? SysAllocString(_p) : 0);
        return p;
    }
};

class CComException: public EXCEPINFO {
public:
	inline CComException() {
		memset((EXCEPINFO*)this, 0, sizeof(EXCEPINFO));
	}
	inline ~CComException() {
		Clear(true);
	}
	inline void Clear(bool internal = false) {
		if (bstrSource) SysFreeString(bstrSource);
		if (bstrDescription) SysFreeString(bstrDescription);
		if (bstrHelpFile) SysFreeString(bstrHelpFile);
		if (!internal) memset((EXCEPINFO*)this, 0, sizeof(EXCEPINFO));
	}
};

template <typename T = IUnknown>
class CComPtr {
public:
    T *p;
    inline CComPtr() : p(0) {}
    inline CComPtr(T *_p) : p(0) { Attach(_p); }
    inline CComPtr(const CComPtr<T> &ptr) : p(0) { if (ptr.p) Attach(ptr.p); }
    inline ~CComPtr() { Release(); }

    inline void Attach(T *_p) { Release(); p = _p; if (p) p->AddRef(); }
    inline T *Detach() { T *pp = p; p = 0; return pp; }
    inline void Release() { if (p) { p->Release(); p = 0; } }

    inline operator T*() const { return p; }
    inline T* operator->() const { return p; }
    inline T& operator*() const { return *p; }
    inline T** operator&() { return &p; }
    inline bool operator!() const { return (p == 0); }
    inline bool operator!=(T* _p) const { return !operator==(_p); }
    inline bool operator==(T* _p) const { return p == _p; }
    inline T* operator = (T* _p) {
        if (p != _p) Attach(_p);
        return p;
    }

    inline HRESULT CoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter = NULL, DWORD dwClsContext = CLSCTX_ALL) {
        Release();
        return ::CoCreateInstance(rclsid, pUnkOuter, dwClsContext, __uuidof(T), (void**)&p);
    }

    inline HRESULT CoCreateInstance(LPCOLESTR szProgID, LPUNKNOWN pUnkOuter = NULL, DWORD dwClsContext = CLSCTX_ALL) {
        Release();
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(szProgID, &clsid);
        if FAILED(hr) return hr;
        return ::CoCreateInstance(clsid, pUnkOuter, dwClsContext, __uuidof(T), (void**)&p);
    }
};


//-------------------------------------------------------------------------------------------------------

template<typename IBASE = IUnknown>
class UnknownImpl : public IBASE {
public:
	inline UnknownImpl() : refcnt(0) {}
	virtual ~UnknownImpl() {}

	// IUnknown interface
	virtual HRESULT __stdcall QueryInterface(REFIID qiid, void **ppvObject) {
		if ((qiid == IID_IUnknown) || (qiid == __uuidof(IBASE))) {
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG __stdcall AddRef() {
		return InterlockedIncrement(&refcnt);
	}

	virtual ULONG __stdcall Release() {
		if (InterlockedDecrement(&refcnt) != 0) return refcnt;
		delete this;
		return 0;
	}

protected:
	LONG refcnt;

};


class DispArrayImpl : public UnknownImpl<IDispatch> {
public:
	CComVariant var;
	DispArrayImpl(const VARIANT &v): var(v) {}

	// IDispatch interface
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};



class DispEnumImpl : public UnknownImpl<IDispatch> {
public:
    CComPtr<IEnumVARIANT> ptr;
    DispEnumImpl() {}
    DispEnumImpl(IEnumVARIANT *p) : ptr(p) {}

    // IDispatch interface
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};


// {9DCE8520-2EFE-48C0-A0DC-951B291872C0}
extern const GUID CLSID_DispObjectImpl;

class DispObjectImpl : public UnknownImpl<IDispatch> {
public:
	Persistent<Object> obj;

	struct name_t { 
		DISPID dispid;
		std::wstring name;
		inline name_t(DISPID id, const std::wstring &nm): dispid(id), name(nm) {}
	};
	typedef std::shared_ptr<name_t> name_ptr;
	typedef std::map<std::wstring, name_ptr> names_t;
	typedef std::map<DISPID, name_ptr> index_t;
	DISPID dispid_next;
	names_t names;
	index_t index;

	inline DispObjectImpl(const Local<Object> &_obj) : obj(Isolate::GetCurrent(), _obj), dispid_next(1) {}
	virtual ~DispObjectImpl() { obj.Reset(); }

	// IUnknown interface
	virtual HRESULT __stdcall QueryInterface(REFIID qiid, void **ppvObject) {
		if (qiid == CLSID_DispObjectImpl) { *ppvObject = this; AddRef(); return S_OK; }
		return UnknownImpl<IDispatch>::QueryInterface(qiid, ppvObject);
	}

	// IDispatch interface
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) { *pctinfo = 0; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { return E_NOTIMPL; }
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
};
