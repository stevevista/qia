#pragma once
#include "unknown_objects.h"
#include "utils.h"
#include <string>
#include <map>


class DispatchCallback;

struct CalllbackInfo {
	DispatchCallback* caller;
	std::wstring name;
	HRESULT hrcode;
	DISPPARAMS *pDispParams;
	VARIANT *pVarResult;
	LONG refernce;
};

class ThreadedCallbackInvokation;

class DispatchCallback : public UnknownImpl<IDispatch> {
public:
    DispatchCallback(const Local<Object> &_obj, CLSID clsid, ITypeInfo* TypeInfo);
	virtual ~DispatchCallback();

	HRESULT initHResult() const;

	// IUnknown interface
	virtual HRESULT __stdcall QueryInterface(REFIID qiid, void **ppvObject);

	// IDispatch interface
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	static bool inMainThread();
	static void dispatchToV8(CalllbackInfo *self, bool dispatched);
	static void WatcherCallback(uv_async_t *w, int revents);

public:
	static NAN_MODULE_INIT(Initialize);

private:
	static DWORD g_threadID;
	static uv_mutex_t    g_queue_mutex;
    static std::queue<ThreadedCallbackInvokation *> g_queue;
    static uv_async_t         g_async;

private:

    HRESULT loadMyTypeInfo();
	std::wstring getNameOfDispId(DISPID);

	HRESULT hrInit_;
	Persistent<Object> obj_;
	ITypeInfo* TypeInfo_;
	CLSID clsid_;
	CLSID clsidTypelib_;
    std::map<DISPID, std::wstring> callbackNames_;
	LONG* pRefernce_;
};

class ThreadedCallbackInvokation {
public:
	ThreadedCallbackInvokation(CalllbackInfo *cbinfo);
    ~ThreadedCallbackInvokation();

    void SignalDoneExecuting();
	void WaitForExecution();

	CalllbackInfo *m_cbinfo;

private:
    uv_cond_t m_cond;
    uv_mutex_t m_mutex;
};
