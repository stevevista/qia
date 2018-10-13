#include "dispatch_callback.h"
using Nan::Get;
using Nan::New;
using Nan::Has;

DWORD DispatchCallback::g_threadID;
uv_mutex_t    DispatchCallback::g_queue_mutex;
std::queue<ThreadedCallbackInvokation *> DispatchCallback::g_queue;
uv_async_t         DispatchCallback::g_async;

// DispatchCallback implementation
DispatchCallback::DispatchCallback(const Local<Object> &obj, CLSID clsid, ITypeInfo* TypeInfo)
: obj_(Isolate::GetCurrent(), obj)
, clsid_(clsid)
, TypeInfo_(TypeInfo)
, pRefernce_(nullptr) {
	hrInit_ = loadMyTypeInfo();
}

DispatchCallback::~DispatchCallback() {
	if (TypeInfo_) TypeInfo_->Release();
	obj_.Reset();
	if (pRefernce_) *pRefernce_ = 1;
}

bool DispatchCallback::inMainThread() {
	return g_threadID == GetCurrentThreadId();
}

HRESULT DispatchCallback::initHResult() const {
	return hrInit_;
}

HRESULT DispatchCallback::loadMyTypeInfo() {
	Nan::HandleScope scope;
	HRESULT hr;

	auto self = obj_.Get(Isolate::GetCurrent());

	// get function dispatch id
	std::vector<std::wstring> callbacks;
	auto names = Nan::GetOwnPropertyNames(self).ToLocalChecked();
	for (auto i = (uint32_t)0; i < names->Length(); i++) {
		auto name = Nan::Get(names, i).ToLocalChecked();
		String::Value vname(name);
		auto cname = std::wstring((LPOLESTR)*vname, vname.length());
		if (cname[0] == L'_' || 
			(cname[0] >= L'a' && cname[0] <= L'z') ||
			(cname[0] >= L'A' && cname[0] <= L'Z')) {
			auto prop = Nan::Get(self, name).ToLocalChecked();
			if (prop->IsFunction()) {
				callbacks.push_back(cname);
			}
		}
	}

	for (auto i = (size_t)0; i < callbacks.size(); i++) {
		LPOLESTR PropName[] = { (LPOLESTR)callbacks[i].c_str() };
		DISPID PropertyID;
		hr = GetIDsOfNames(IID_NULL, PropName, 1, LOCALE_SYSTEM_DEFAULT, &PropertyID);
		if (!SUCCEEDED(hr)) {
			Nan::ThrowError(DispError(Isolate::GetCurrent(), hr, PropName[0]));
			return hr;
		}

		callbackNames_.insert(std::make_pair(PropertyID, callbacks[i]));
	}

	return hr;
}

HRESULT __stdcall DispatchCallback::QueryInterface(REFIID qiid, void **ppvObject) {
	if (qiid == clsid_) { 
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	return UnknownImpl<IDispatch>::QueryInterface(qiid, ppvObject);
}

HRESULT STDMETHODCALLTYPE DispatchCallback::GetTypeInfoCount(UINT *pctinfo) { 
	if (!TypeInfo_)
		*pctinfo = 0;
	else 
		*pctinfo = 1; 
	return S_OK; 
}

HRESULT STDMETHODCALLTYPE DispatchCallback::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) { 
	if (!TypeInfo_)
		return E_NOTIMPL;

	*ppTInfo = TypeInfo_;
	TypeInfo_->AddRef();
	return S_OK; 
}

HRESULT STDMETHODCALLTYPE DispatchCallback::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (!TypeInfo_)
		return E_NOTIMPL;

	return(DispGetIDsOfNames(TypeInfo_, rgszNames, cNames, rgDispId));
}

std::wstring DispatchCallback::getNameOfDispId(DISPID dispid) {

	auto it = callbackNames_.find(dispid);
	return it == callbackNames_.end() ? L"" : it->second;
}

HRESULT STDMETHODCALLTYPE DispatchCallback::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	// Prepare name by member id
	std::wstring cname;
	if (dispIdMember != DISPID_VALUE) {
		cname = getNameOfDispId(dispIdMember);
	}

	if (cname.empty()) {
		return DISP_E_MEMBERNOTFOUND;
	}

	// Set property value
	if ((wFlags & DISPATCH_PROPERTYPUT) != 0) {
		return E_NOTIMPL;
	}

	// Get property value
	if ((wFlags & DISPATCH_PROPERTYGET) != 0) {
		return E_NOTIMPL;
	}

	// Call property as method
	if ((wFlags & DISPATCH_METHOD) != 0) {
		CalllbackInfo cbinfo;
		cbinfo.caller = this;
		cbinfo.name = cname;
		cbinfo.hrcode = S_OK;
		cbinfo.pDispParams = pDispParams;
		cbinfo.pVarResult = pVarResult;
		cbinfo.refernce = 0;

		// refernce to as indicator , incase the object is unadvised in another thread
		this->pRefernce_ = &cbinfo.refernce;

		if (inMainThread()) {
			dispatchToV8(&cbinfo, false);
		} else {
			// hold the event loop open while this is executing
			uv_mutex_lock(&g_queue_mutex);
    		uv_ref((uv_handle_t *)&g_async);
			uv_mutex_unlock(&g_queue_mutex);
			// create a temporary storage area for our invokation parameters
    		ThreadedCallbackInvokation inv(&cbinfo);

			// push it to the queue -- threadsafe
			uv_mutex_lock(&g_queue_mutex);
			g_queue.push(&inv);
			uv_mutex_unlock(&g_queue_mutex);

			// send a message to our main thread to wake up the WatchCallback loop
			uv_async_send(&g_async);

			// wait for signal from calling thread
			inv.WaitForExecution();

			uv_mutex_lock(&g_queue_mutex);
			uv_unref((uv_handle_t *)&g_async);
			uv_mutex_unlock(&g_queue_mutex);
		}

		return cbinfo.hrcode;
	}

	return S_OK;
}


NAN_MODULE_INIT(DispatchCallback::Initialize) {

	// initialize our threaded invokation stuff
  	g_threadID = GetCurrentThreadId();

	uv_async_init(uv_default_loop(), &g_async, (uv_async_cb) WatcherCallback);
	uv_mutex_init(&g_queue_mutex);

	// allow the event loop to exit while this is running
	uv_unref((uv_handle_t *)&g_async);
}

void DispatchCallback::dispatchToV8(CalllbackInfo *self, bool dispatched) {

	if (self->refernce != 0) {
		// this object has been unadvised
		return;
	}

	self->caller->pRefernce_ = nullptr;

	Nan::HandleScope scope;

	auto isolate = Isolate::GetCurrent();
	NodeArguments args(isolate, self->pDispParams, true);
	int argcnt = (int)args.items.size();
	Local<Value> *argptr = (argcnt > 0) ? &args.items[0] : nullptr;
	Local<Object> object = self->caller->obj_.Get(isolate);

	auto name = New<String>((uint16_t*)self->name.c_str()).ToLocalChecked();
	auto val = object->Get(name);

	Local<Value> ret;

	if (val->IsFunction()) {
		Local<Function> func = Local<Function>::Cast(val);
		if (!func.IsEmpty()) {
			ret = func->Call(object, argcnt, argptr);
		} else {
			if (dispatched) {
				Nan::ThrowError("fatal: callback may be garbage collected!");
			}
			self->hrcode = DISP_E_BADCALLEE;
		}
	} else {
		if (dispatched) {
			Nan::ThrowError("fatal: callback may be garbage collected!");
		}
		self->hrcode = DISP_E_BADCALLEE;
	}

	// Store result
	if (self->pVarResult) {
		Value2Variant(isolate, ret, *self->pVarResult, VT_NULL);
	}
}

void DispatchCallback::WatcherCallback(uv_async_t *w, int revents) {
	uv_mutex_lock(&g_queue_mutex);

	while (!g_queue.empty()) {
		ThreadedCallbackInvokation *inv = g_queue.front();
		g_queue.pop();

		dispatchToV8(inv->m_cbinfo, true);
		inv->SignalDoneExecuting();
	}

	uv_mutex_unlock(&g_queue_mutex);
}


/////////////////////////////////////////////////////////////////////////////////////////
void ThreadedCallbackInvokation::WaitForExecution() {
  	uv_cond_wait(&m_cond, &m_mutex);
}

ThreadedCallbackInvokation::ThreadedCallbackInvokation(CalllbackInfo *cbinfo) {
	m_cbinfo = cbinfo;

  	uv_mutex_init(&m_mutex);
  	uv_mutex_lock(&m_mutex);
  	uv_cond_init(&m_cond);
}

ThreadedCallbackInvokation::~ThreadedCallbackInvokation() {
  uv_mutex_unlock(&m_mutex);
  uv_cond_destroy(&m_cond);
  uv_mutex_destroy(&m_mutex);
}

void ThreadedCallbackInvokation::SignalDoneExecuting() {
  uv_mutex_lock(&m_mutex);
  uv_cond_signal(&m_cond);
  uv_mutex_unlock(&m_mutex);
}
