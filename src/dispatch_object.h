#pragma once
#include "disp.h"

class DispObject: public ObjectWrap
{
public:
	DispObject(const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1, int opt = 0, const std::wstring& inprocServer32 = L"");
	~DispObject();

	static Persistent<ObjectTemplate> inst_template;
	static Persistent<FunctionTemplate> clazz_template;
	static void NodeInit(const Local<Object> &target);
	static bool HasInstance(Isolate *isolate, const Local<Value> &obj) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		return !clazz.IsEmpty() && clazz->HasInstance(obj);
	}
	static bool GetValueOf(Isolate *isolate, const Local<Object> &obj, VARIANT &value) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty() || !clazz->HasInstance(obj)) return false;
		DispObject *self = Unwrap<DispObject>(obj);
		return self && SUCCEEDED(self->valueOf(isolate, value));
	}
	static Local<Object> NodeCreate(Isolate *isolate, IDispatch *disp, const std::wstring &name, int opt) {
		Local<Object> parent;
		DispInfoPtr ptr(new DispInfo(disp, name, opt));
		return DispObject::NodeCreate(isolate, parent, ptr, name);
	}

private:
	static Local<Object> NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, const std::wstring &name, DISPID id = DISPID_UNKNOWN, LONG indx = -1, int opt = 0, const std::wstring& inprocServer32 = L"");

	static void NodeCreate(const FunctionCallbackInfo<Value> &args);
	static void NodeValueOf(const FunctionCallbackInfo<Value> &args);
	static void NodeToString(const FunctionCallbackInfo<Value> &args);
	static void NodeRelease(const FunctionCallbackInfo<Value> &args);
	static void NodeCast(const FunctionCallbackInfo<Value> &args);
    static void NodeGet(Local<String> name, const PropertyCallbackInfo<Value> &args);
	static void NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static void NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value> &args);
	static void NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value> &args);
	static NAN_METHOD(NodeCall);
	static NAN_METHOD(ConnectionAdvise);
	static NAN_METHOD(ConnectionUnadvise);

protected:
	bool release();
	bool get(LPOLESTR tag, LONG index, const PropertyCallbackInfo<Value> &args);
	bool set(LPOLESTR tag, LONG index, const Local<Value> &value, const PropertyCallbackInfo<Value> &args);
	void call(Isolate *isolate, const Nan::FunctionCallbackInfo<Value> &args);

	HRESULT valueOf(Isolate *isolate, VARIANT &value);
	HRESULT valueOf(Isolate *isolate, const Local<Object> &self, Local<Value> &value);
	void toString(const FunctionCallbackInfo<Value> &args);
    Local<Value> getIdentity(Isolate *isolate);
    Local<Value> getTypeInfo(Isolate *isolate);

private:
	int options;
	inline bool is_null() { return !disp; }
	inline bool is_prepared() { return (options & option_prepared) != 0; }
	inline bool is_object() { return dispid == DISPID_VALUE /*&& index < 0*/; }
	inline bool is_owned() { return (options & option_owned) != 0; }

	static void AsyncCall(uv_work_t *req);
    static void FinishAsyncCall(uv_work_t *req);
	class DispWorker;
	
	DispInfoPtr disp;
	std::wstring name;
	DISPID dispid;
	LONG index;

	struct connection_t {
		CLSID clsid;
		CComPtr<IUnknown> unk;
	};
	std::map<DWORD, connection_t> connections_;

	LPTYPELIB pTypelib_;
	std::wstring inprocServer32_;
	static bool is64arch;

	HRESULT prepare();
};

