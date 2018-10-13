#include "dispatch_object.h"
#include "dispatch_callback.h"

using Nan::ThrowError;
using Nan::TypeError;
using Nan::Get;
using Nan::New;
using Nan::AsyncQueueWorker;
using Nan::AsyncWorker;

Persistent<ObjectTemplate> DispObject::inst_template;
Persistent<FunctionTemplate> DispObject::clazz_template;

bool DispObject::is64arch = false;


DispObject::DispObject(const DispInfoPtr &ptr, const std::wstring &nm, DISPID id, LONG indx, int opt, const std::wstring& inprocServer32)
	: disp(ptr), options((ptr->options & option_mask) | opt), name(nm), dispid(id), index(indx)
	, inprocServer32_(inprocServer32)
	, pTypelib_(nullptr)
{	
	if (dispid == DISPID_UNKNOWN) {
		dispid = DISPID_VALUE;
        options |= option_prepared;
	}
	else options |= option_owned;
	NODE_DEBUG_FMT("DispObject '%S' constructor", name.c_str());
}

DispObject::~DispObject() {
	if (pTypelib_) pTypelib_->Release();
	NODE_DEBUG_FMT("DispObject '%S' destructor", name.c_str());
}

HRESULT DispObject::prepare() {
	CComVariant value;
	HRESULT hrcode = disp ? disp->GetProperty(dispid, index, &value) : E_UNEXPECTED;

	// Init dispatch interface
	options |= option_prepared;
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&value, &ptr)) {
		disp.reset(new DispInfo(ptr, name, options, &disp));
		dispid = DISPID_VALUE;
	}
	else if ((value.vt & VT_ARRAY) != 0) {
		
	}
	return hrcode;
}

bool DispObject::release() {
    if (!disp) return false;
    NODE_DEBUG_FMT("DispObject '%S' release", name.c_str());
    disp.reset();
    return true;
}


bool DispObject::get(LPOLESTR tag, LONG index, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	if (!is_prepared()) prepare();
    if (!disp) {
        isolate->ThrowException(DispErrorNull(isolate));
        return false;
    }

	// Search dispid
    HRESULT hrcode;
    DISPID propid;
	bool prop_by_key = false;
    if (!tag) {
        tag = (LPOLESTR)name.c_str();
        propid = dispid;
    }
	else {
        hrcode = disp->FindProperty(tag, &propid);
        if (SUCCEEDED(hrcode) && propid == DISPID_UNKNOWN) hrcode = E_INVALIDARG;
        if FAILED(hrcode) {
			prop_by_key = (options & option_property) != 0;
			if (!prop_by_key) {
				//isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyFind", tag));
				args.GetReturnValue().SetUndefined();
				return false;
			}
			propid = dispid;
        }
    }

	// Check type info
	int opt = 0;
	bool is_property_simple = false;
	if (prop_by_key) {
		is_property_simple = true;
		opt |= option_property;
	}
	else {
		DispInfo::type_ptr disp_info;
		if (disp->GetTypeInfo(propid, disp_info)) {
			if (disp_info->is_function_simple()) opt |= option_function_simple;
			else {
				if (disp_info->is_property()) opt |= option_property;
				is_property_simple = disp_info->is_property_simple();
			}
		}
	}

    // Return as property value
	if (is_property_simple) {
		CComException except;
		CComVariant value;
		VarArguments vargs;
		if (prop_by_key) vargs.items.push_back(CComVariant(tag));
		if (index >= 0) vargs.items.push_back(CComVariant(index));
		LONG argcnt = (LONG)vargs.items.size();
		VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
		hrcode = disp->GetProperty(propid, argcnt, pargs, &value, &except);
		if (FAILED(hrcode) && dispid != DISPID_VALUE){
			isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyGet", tag, &except));
			return false;
		}
		CComPtr<IDispatch> ptr;
		if (VariantDispGet(&value, &ptr)) {
			DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
			Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp_result, tag, DISPID_UNKNOWN, -1, opt);
			args.GetReturnValue().Set(result);
		}
		else {
			args.GetReturnValue().Set(Variant2Value(isolate, value));
		}
	}

	// Return as dispatch object 
	else {
		Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp, tag, propid, index, opt, inprocServer32_);
		args.GetReturnValue().Set(result);
	}
	return true;
}

bool DispObject::set(LPOLESTR tag, LONG index, const Local<Value> &value, const PropertyCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	if (!is_prepared()) prepare();
    if (!disp) {
        isolate->ThrowException(DispErrorNull(isolate));
        return false;
    }
	
	// Search dispid
	HRESULT hrcode;
	DISPID propid;
	if (!tag) {
		tag = (LPOLESTR)name.c_str();
		propid = dispid;
	}
	else {
		hrcode = disp->FindProperty(tag, &propid);
		if (SUCCEEDED(hrcode) && propid == DISPID_UNKNOWN) hrcode = E_INVALIDARG;
		if FAILED(hrcode) {
			isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyFind", tag));
			return false;
		}
	}

	// Set value using dispatch
	CComException except;
    CComVariant ret;
	VarArguments vargs(isolate, value);
	if (index >= 0) vargs.items.push_back(CComVariant(index));
	LONG argcnt = (LONG)vargs.items.size();
    VARIANT *pargs = (argcnt > 0) ? &vargs.items.front() : 0;
	hrcode = disp->SetProperty(propid, argcnt, pargs, &ret, &except);
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"DispPropertyPut", tag, &except));
        return false;
    }

	// Send result
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&ret, &ptr)) {
		std::wstring rtag;
		rtag.reserve(32);
		rtag += L"@";
		rtag += tag;
		DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
		Local<Object> result = DispObject::NodeCreate(isolate, args.This(), disp_result, rtag);
		args.GetReturnValue().Set(result);
	}
	else {
		args.GetReturnValue().Set(Variant2Value(isolate, ret));
	}
    return true;
}

HRESULT DispObject::valueOf(Isolate *isolate, VARIANT &value) {
	if (!is_prepared()) prepare();
	HRESULT hrcode;
	if (!disp) hrcode = E_UNEXPECTED;

	// simple function without arguments
	else if ((options & option_function_simple) != 0) {
		hrcode = disp->ExecuteMethod(dispid, 0, 0, &value);
	}

	// property or array element
	else if (dispid != DISPID_VALUE || index >= 0) {
		hrcode = disp->GetProperty(dispid, index, &value);
	}

	// self dispatch object
	else /*if (is_object())*/ {
		value.vt = VT_DISPATCH;
		value.pdispVal = (IDispatch*)disp->ptr;
		if (value.pdispVal) value.pdispVal->AddRef();
		hrcode = S_OK;
	}
	return hrcode;
}

HRESULT DispObject::valueOf(Isolate *isolate, const Local<Object> &self, Local<Value> &value) {
	if (!is_prepared()) prepare();
	HRESULT hrcode;
	if (!disp) hrcode = E_UNEXPECTED;
	else {
		CComVariant val;

		// simple function without arguments
		if ((options & option_function_simple) != 0) {
			hrcode = disp->ExecuteMethod(dispid, 0, 0, &val);
		}

		// self value, property or array element
		else {
			hrcode = disp->GetProperty(dispid, index, &val);
		}

		// convert result to v8 value
		if SUCCEEDED(hrcode) {
			value = Variant2Value(isolate, val);
		}

		// or return self as object
		else  {
			value = self;
			hrcode = S_OK;
		}
	}
	return hrcode;
}

void DispObject::toString(const FunctionCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	CComVariant val;
	HRESULT hrcode = valueOf(isolate, val);
	if FAILED(hrcode) {
		isolate->ThrowException(Win32Error(isolate, hrcode, L"DispToString"));
		return;
	}
	args.GetReturnValue().Set(Variant2String(isolate, val));
}

Local<Value> DispObject::getIdentity(Isolate *isolate) {
    std::wstring id;
    id.reserve(128);
    id += name;
    DispInfoPtr ptr = disp;
    if (ptr && ptr->name == id)
        ptr = ptr->parent.lock();
    while (ptr) {
        id.insert(0, L".");
        id.insert(0, ptr->name);
        ptr = ptr->parent.lock();
    }
    return String::NewFromTwoByte(isolate, (uint16_t*)id.c_str());
}

Local<Value> DispObject::getTypeInfo(Isolate *isolate) {
    if ((options & option_type) == 0 || !disp) {
        return Undefined(isolate);
    }
    uint32_t index = 0;
    Local<v8::Array> items(v8::Array::New(isolate));
    disp->Enumerate([isolate, this, &items, &index](ITypeInfo *info, FUNCDESC *desc) {
        CComBSTR name;
        this->disp->GetItemName(info, desc->memid, &name); 
        Local<Object> item(Object::New(isolate));
        if (name) item->Set(String::NewFromUtf8(isolate, "name"), String::NewFromTwoByte(isolate, (uint16_t*)(BSTR)name));
        item->Set(String::NewFromUtf8(isolate, "dispid"), Int32::New(isolate, desc->memid));
        item->Set(String::NewFromUtf8(isolate, "invkind"), Int32::New(isolate, desc->invkind));
        item->Set(String::NewFromUtf8(isolate, "argcnt"), Int32::New(isolate, desc->cParams));
        items->Set(index++, item);
    });
    return items;
}

//-----------------------------------------------------------------------------------
// Static Node JS callbacks

void DispObject::NodeInit(const Local<Object> &target) {

	SYSTEM_INFO si;
	GetNativeSystemInfo(&si);
	if (si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64 ||
		si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_IA64 ) {
		is64arch = true;
	}

    Isolate *isolate = target->GetIsolate();

    // Prepare constructor template
    Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
    clazz->SetClassName(String::NewFromUtf8(isolate, "Dispatch"));

	NODE_SET_PROTOTYPE_METHOD(clazz, "toString", NodeToString);
	NODE_SET_PROTOTYPE_METHOD(clazz, "valueOf", NodeValueOf);

    Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
    inst->SetInternalFieldCount(1);
    inst->SetNamedPropertyHandler(NodeGet, NodeSet);
    inst->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex);
    Nan::SetCallAsFunctionHandler(inst, NodeCall);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__id"), NodeGet);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__value"), NodeGet);
    inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__type"), NodeGet);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__inprocServer32"), NodeGet);

    inst_template.Reset(isolate, inst);
	clazz_template.Reset(isolate, clazz);
    target->Set(String::NewFromUtf8(isolate, "Object"), clazz->GetFunction());
	target->Set(String::NewFromUtf8(isolate, "cast"), FunctionTemplate::New(isolate, NodeCast, target)->GetFunction());
	target->Set(String::NewFromUtf8(isolate, "release"), FunctionTemplate::New(isolate, NodeRelease, target)->GetFunction());

    //Context::GetCurrent()->Global()->Set(String::NewFromUtf8("ActiveXObject"), t->GetFunction());
	NODE_DEBUG_MSG("DispObject initialized");
}

Local<Object> DispObject::NodeCreate(Isolate *isolate, const Local<Object> &parent, const DispInfoPtr &ptr, const std::wstring &name, DISPID id, LONG index, int opt, const std::wstring& inprocServer32) {
    Local<Object> self;
    if (!inst_template.IsEmpty()) {
        self = inst_template.Get(isolate)->NewInstance();
        (new DispObject(ptr, name, id, index, opt, inprocServer32))->Wrap(self);
		//Local<String> prop_id(String::NewFromUtf8(isolate, "_identity"));
		//self->Set(prop_id, String::NewFromTwoByte(isolate, (uint16_t*)name));
	}
    return self;
}

void DispObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
    Isolate *isolate = args.GetIsolate();
    int argcnt = args.Length();
    if (argcnt < 1) {
        isolate->ThrowException(InvalidArgumentsError(isolate));
        return;
    }

	std::wstring inprocServer32;

    int options = option_type;
    if (argcnt > 1) {
        Local<Value> argopt = args[1];
        if (!argopt.IsEmpty() && argopt->IsObject()) {
            Local<Object> opt = argopt->ToObject();
            if (!v8val2bool(opt->Get(String::NewFromUtf8(isolate, "type")), true)) {
                options &= ~option_type;
            }
		}
    }
    
    // Invoked as plain function
    if (!args.IsConstructCall()) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty()) {
			isolate->ThrowException(TypeError(isolate, "FunctionTemplateIsEmpty"));
			return;
		}
        const int argc = 1;
        Local<Value> argv[argc] = { args[0] };
        Local<Context> context = isolate->GetCurrentContext();
        Local<Function> cons = Local<Function>::New(isolate, clazz->GetFunction());
        Local<Object> self = cons->NewInstance(context, argc, argv).ToLocalChecked();
        args.GetReturnValue().Set(self);
        return;
    }

	// Create dispatch object from ProgId
	HRESULT hrcode;
	std::wstring name;
	CComPtr<IDispatch> disp;
	if (args[0]->IsString()) {

		// Prepare arguments
		Local<String> progid = args[0]->ToString();
		String::Value vname(progid);
		if (vname.length() <= 0) hrcode = E_INVALIDARG;
		else {
			name.assign((LPOLESTR)*vname, vname.length());
			CLSID clsid;
			hrcode = CLSIDFromProgID(name.c_str(), &clsid);
			if SUCCEEDED(hrcode) {

				wchar_t clsid_buf[128];
				wchar_t regkey[128];
				StringFromGUID2(clsid, clsid_buf, 128);

				// get InprocServer32
				if (is64arch) {
					#ifdef _WIN64
					wsprintfW(regkey, L"CLSID\\%s\\InprocServer32", clsid_buf);
					#else
					wsprintfW(regkey, L"WOW6432Node\\CLSID\\%s\\InprocServer32", clsid_buf);
					#endif
				} else {
					wsprintfW(regkey, L"CLSID\\%s\\InprocServer32", clsid_buf);
				}

				HKEY hKey;
				DWORD dwType = REG_SZ; 
				DWORD dwSize = MAX_PATH;
				wchar_t data[MAX_PATH];

				if (RegOpenKeyW(HKEY_CLASSES_ROOT, regkey, &hKey) == ERROR_SUCCESS) {
					if (RegQueryValueExW(hKey, L"CodeBase", 
							nullptr, 
							&dwType, 
							(LPBYTE)data, 
							&dwSize) == ERROR_SUCCESS) {
						inprocServer32 = data;

					} else {
						dwSize = MAX_PATH;
						if (RegQueryValueExW(hKey, nullptr, 
								nullptr, 
								&dwType, 
								(LPBYTE)data, 
								&dwSize) == ERROR_SUCCESS) {
							inprocServer32 = data;
						}
					}

					if (inprocServer32.find(L"file:///") == 0) {
						inprocServer32 = inprocServer32.substr(8);
						for (auto& c : inprocServer32) {
							if (c == L'/') c = L'\\';
						}
					}

					RegCloseKey(hKey);
				}

				hrcode = disp.CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER);
			}
		}
	}

	// Create dispatch object from javascript object
	else if (args[0]->IsObject()) {
		name = L"#";
		disp = new DispObjectImpl(args[0]->ToObject());
		hrcode = S_OK;
	}

	// Other
	else {
		hrcode = E_INVALIDARG;
	}

	// Prepare result
	if FAILED(hrcode) {
		isolate->ThrowException(DispError(isolate, hrcode, L"CreateInstance", name.c_str()));
	}
	else {
		Local<Object> &self = args.This();
		DispInfoPtr ptr(new DispInfo(disp, name, options));
		(new DispObject(ptr, name, DISPID_UNKNOWN, -1, 0, inprocServer32))->Wrap(self);
		args.GetReturnValue().Set(self);
	}
}

void DispObject::NodeGet(Local<String> name, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
    NODE_DEBUG_FMT2("DispObject '%S.%S' get", self->name.c_str(), id);
    if (_wcsicmp(id, L"__value") == 0) {
        Local<Value> result;
        HRESULT hrcode = self->valueOf(isolate, args.This(), result);
        if FAILED(hrcode) isolate->ThrowException(Win32Error(isolate, hrcode, L"DispValueOf"));
        else args.GetReturnValue().Set(result);
    }
    else if (_wcsicmp(id, L"__id") == 0) {
		args.GetReturnValue().Set(self->getIdentity(isolate));
	}
    else if (_wcsicmp(id, L"__type") == 0) {
        args.GetReturnValue().Set(self->getTypeInfo(isolate));
    }
	else if (_wcsicmp(id, L"__proto__") == 0) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty()) args.GetReturnValue().SetNull();
		else args.GetReturnValue().Set(clazz_template.Get(isolate)->GetFunction());
	}
	else if (_wcsicmp(id, L"valueOf") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeValueOf, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"toString") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeToString, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"callbackAdvise") == 0) {
		args.GetReturnValue().Set(Nan::New<FunctionTemplate>(ConnectionAdvise, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"callbackUnadvise") == 0) {
		args.GetReturnValue().Set(Nan::New<FunctionTemplate>(ConnectionUnadvise, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"__inprocServer32") == 0) {
		args.GetReturnValue().Set(String::NewFromTwoByte(isolate, (uint16_t*)self->inprocServer32_.c_str()));
	}
	else {
		self->get(id, -1, args);
	}
}

void DispObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
    NODE_DEBUG_FMT2("DispObject '%S[%u]' get", self->name.c_str(), index);
    self->get(0, index, args);
}

void DispObject::NodeSet(Local<String> name, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
	NODE_DEBUG_FMT2("DispObject '%S.%S' set", self->name.c_str(), id);
    self->set(id, -1, value, args);
}

void DispObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
    Isolate *isolate = args.GetIsolate();
    DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	NODE_DEBUG_FMT2("DispObject '%S[%u]' set", self->name.c_str(), index);
	self->set(0, index, value, args);
}

NAN_METHOD(DispObject::NodeCall) {
    Isolate *isolate = Isolate::GetCurrent();
    DispObject *self = DispObject::Unwrap<DispObject>(info.This());
	if (!self) {
		Nan::ThrowError(DispErrorInvalid(isolate));
		return;
	}
	NODE_DEBUG_FMT("DispObject '%S' call", self->name.c_str());
    self->call(isolate, info);
}

void DispObject::NodeValueOf(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result;
	HRESULT hrcode = self->valueOf(isolate, args.This(), result);
	if FAILED(hrcode) {
		isolate->ThrowException(Win32Error(isolate, hrcode, L"DispValueOf"));
		return;
	}
	args.GetReturnValue().Set(result);
}

void DispObject::NodeToString(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	DispObject *self = DispObject::Unwrap<DispObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	self->toString(args);
}

void DispObject::NodeRelease(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
    int rcnt = 0, argcnt = args.Length();
    for (int argi = 0; argi < argcnt; argi++) {
        Local<Value> &obj = args[argi];
        if (obj->IsObject()) {
            Local<Object> disp_obj = obj->ToObject();
            DispObject *disp = DispObject::Unwrap<DispObject>(disp_obj);
            if (disp && disp->release())
                rcnt ++;
        }
    }
    args.GetReturnValue().Set(rcnt);
}

void DispObject::NodeCast(const FunctionCallbackInfo<Value>& args) {
	Local<Object> inst = VariantObject::NodeCreateInstance(args);
	args.GetReturnValue().Set(inst);
}


class DispObject::DispWorker : public AsyncWorker {
public:
	DispWorker(const Nan::FunctionCallbackInfo<Value> &info, DispObject* ptr)
    : AsyncWorker(new Nan::Callback(Nan::To<Function>(info[info.Length() - 1]).ToLocalChecked()))
	, self(ptr) {
		Nan::HandleScope scope;

		SaveToPersistent("parent", info.This());
		
		const int argsCount = info.Length() - 1;
		args.resize(argsCount);
		for (int i = 0; i < argsCount; i ++) {
			Value2Variant(Isolate::GetCurrent(), info[argsCount - i - 1], args[i]);
		}
	}

	void Execute() {
		LONG argsCount = args.size();
		VARIANT *pargs = (argsCount > 0) ? &args.front() : 0;
		hrcode = self->disp->ExecuteMethod(self->dispid, argsCount, pargs, &ret, &except);
		if (FAILED(hrcode)) {
			SetErrorMessage("error");
		}
	}

	void HandleOKCallback() {
		Nan::HandleScope scope;

		// Prepare result
        Local<Value> result;
        CComPtr<IDispatch> ptr;
        if (VariantDispGet(&ret, &ptr)) {
            std::wstring tag;
            tag.reserve(32);
            tag += L"@";
            tag += self->name;
            DispInfoPtr disp_result(new DispInfo(ptr, tag, self->options, &self->disp));
			Local<Value> parent = GetFromPersistent("parent");
            result = DispObject::NodeCreate(Isolate::GetCurrent(), parent->ToObject(), disp_result, tag, DISPID_UNKNOWN, -1, 0, self->inprocServer32_);
        }
        else {
            result = Variant2Value(Isolate::GetCurrent(), ret);
        }

		Local<Value> argv[] = {
			Nan::Null(), result
		};

		callback->Call(2, argv, async_resource);
	}

	void HandleErrorCallback() {
		Nan::HandleScope scope;

		Local<Value> argv[] = {
			DispError(Isolate::GetCurrent(), hrcode, L"DispInvoke", self->name.c_str(), &except)
		};
		callback->Call(1, argv, async_resource);
	}

private:
	std::vector<CComVariant> args;
	DispObject* self;
	CComVariant ret;
	HRESULT hrcode;
    CComException except;
};

void DispObject::call(Isolate *isolate, const Nan::FunctionCallbackInfo<Value> &args) {
    Nan::HandleScope scope;
    if (!disp) {
        ThrowError(DispErrorNull(isolate));
        return;
    }

    auto argsCount = args.Length();
    if (!(options & option_property) &&
			argsCount > 0 && args[argsCount - 1]->IsFunction()) {

		AsyncQueueWorker(new DispWorker(args, this));
        args.GetReturnValue().SetUndefined();
        return;
    }

    CComException except;
	CComVariant ret;
	LONG argcnt = args.Length();
	std::vector<CComVariant> vargs(argcnt);
	for (int i = 0; i < argcnt; i ++)
		Value2Variant(isolate, args[argcnt - i - 1], vargs[i]);

	VARIANT *pargs = (argcnt > 0) ? &vargs.front() : 0;
	HRESULT hrcode;

	if ((options & option_property) == 0) hrcode = disp->ExecuteMethod(dispid, argcnt, pargs, &ret, &except);
	else hrcode = disp->GetProperty(dispid, argcnt, pargs, &ret, &except);
    if FAILED(hrcode) {
        ThrowError(DispError(isolate, hrcode, L"DispInvoke", name.c_str(), &except));
        return;
    }

	// Prepare result
    Local<Value> result;
	CComPtr<IDispatch> ptr;
	if (VariantDispGet(&ret, &ptr)) {
        std::wstring tag;
        tag.reserve(32);
        tag += L"@";
        tag += name;
		DispInfoPtr disp_result(new DispInfo(ptr, tag, options, &disp));
		result = DispObject::NodeCreate(isolate, args.This(), disp_result, tag);
	}
	else {
		result = Variant2Value(isolate, ret);
	}
    args.GetReturnValue().Set(result);
}


NAN_METHOD(DispObject::ConnectionAdvise) {

    if (info.Length() < 1 || info[0].IsEmpty() || !info[0]->IsObject()) {
        ThrowError("callbackAdvise 1'st paramter should be object");
        return;
    }

	auto self = DispObject::Unwrap<DispObject>(info.This());
	auto obj = info[0]->ToObject();
	Local<String> typeClsid;
	if (Nan::Has(obj, New("__typelib").ToLocalChecked()).FromJust()) {
		typeClsid = Get(obj, New("__typelib").ToLocalChecked()).ToLocalChecked()->ToString();
	}

	// load typelib
	LPTYPELIB pTypeLib = nullptr;

	if (!typeClsid.IsEmpty()) {
		CLSID clsid;
		String::Value vname(typeClsid);
		auto hr = CLSIDFromString((LPCOLESTR)*vname, &clsid);

		if (!SUCCEEDED(hr)) {
			Nan::ThrowError(DispError(Isolate::GetCurrent(), hr, L"GetConnectionTypeInfo", (LPCOLESTR)*vname));
			return;
		}

		if (FAILED(hr = LoadRegTypeLib(clsid, 1, 0, 0, &pTypeLib))) {
			Nan::ThrowError(DispError(Isolate::GetCurrent(), hr, L"LoadRegTypeLib", (LPCOLESTR)*vname));
			return;
		}
	}

	if (!pTypeLib) {
		if (self->pTypelib_) {
			self->pTypelib_->AddRef();
			pTypeLib = self->pTypelib_;
		} else {
			if (!self->inprocServer32_.empty()) {
				auto code = LoadTypeLibEx(self->inprocServer32_.c_str(),
					REGKIND_NONE,
					&pTypeLib
					);
				if (FAILED(code)) {
					auto tlbfile = self->inprocServer32_.substr(0, self->inprocServer32_.size()-3);
					tlbfile = tlbfile + L"tlb";
					code = LoadTypeLibEx(tlbfile.c_str(), REGKIND_NONE, &pTypeLib);
				}

				if (SUCCEEDED(code)) {
					pTypeLib->AddRef();
					self->pTypelib_ = pTypeLib;
				}
			}
		}
	}

    auto interStr = Get(obj, New("__interface").ToLocalChecked()).ToLocalChecked()->ToString();
    String::Value vinterface(interStr);

	wchar_t szNameBuf[128] = {0};
	if (!interStr->IsUndefined()) {
		wcscpy(szNameBuf, (const wchar_t*)*vinterface);
	}

	CLSID clsid;
	ITypeInfo* pTypeinfo = nullptr;
	ITypeInfo* pTypeinfos[5];
	MEMBERID memid[5];
	unsigned short found = 5;
	auto hr = pTypeLib->FindName(szNameBuf, 0, pTypeinfos, memid, &found);
	pTypeLib->Release();

	if (FAILED(hr)) {
		Nan::ThrowError(DispError(Isolate::GetCurrent(), hr, L"FindName", (LPCOLESTR)*vinterface));
		return;
	}

    for (int i=0; i<found; i++) {
		if (memid[i] == MEMBERID_NIL) {

			TYPEATTR* pTypeAttr;
            hr = pTypeinfos[i]->GetTypeAttr(&pTypeAttr);
            if(FAILED(hr)) break;

			if(pTypeAttr->typekind==TKIND_INTERFACE||
                    pTypeAttr->typekind==TKIND_DISPATCH)
            {
                pTypeinfo = pTypeinfos[i];
				clsid = pTypeAttr->guid;
            }

			pTypeinfos[i]->ReleaseTypeAttr(pTypeAttr);

			if (pTypeinfo)
				break;
		}
	}

	if (FAILED(hr) || !pTypeinfo) {
		Nan::ThrowError(DispError(Isolate::GetCurrent(), hr, L"LoadTypeInfo", (LPCOLESTR)*vinterface));
		return;
	}

    CComPtr<IConnectionPointContainer> cp_cont;

    bool throwed = false;
    hr = self->disp->ptr->QueryInterface(&cp_cont);
    if (SUCCEEDED(hr)) {
		CComPtr<IConnectionPoint> point;
		hr = cp_cont->FindConnectionPoint(clsid, &point);

		if (SUCCEEDED(hr) && point) {
            CComPtr<IUnknown> unk;
            auto callbackObj = new DispatchCallback(obj, clsid, pTypeinfo);
            hr = callbackObj->initHResult();

            if (SUCCEEDED(hr)) {
                unk.Attach(callbackObj);

                DWORD dwCookie;
                hr = point->Advise(unk, &dwCookie);
                if (SUCCEEDED(hr)) {
                    self->connections_.insert(std::make_pair(dwCookie, connection_t{clsid, unk}));
                    info.GetReturnValue().Set(New((uint32_t)dwCookie));
                }
            } else {
                throwed = true;
                callbackObj->Release();
            }
            
            point.Release();
		}
        cp_cont.Release();			
    }

    if (!SUCCEEDED(hr)) {
        if (!throwed)
		    ThrowError(DispError(Isolate::GetCurrent(), hr, L"ConnectionAdvise"));
		return;
	}
}

NAN_METHOD(DispObject::ConnectionUnadvise) {
     if (info.Length() < 1) {
        return;
    }

    auto dwCookie = info[0]->Uint32Value();

    auto self = DispObject::Unwrap<DispObject>(info.This());
	auto it = self->connections_.find(dwCookie);
	if (it != self->connections_.end()) {
		CComPtr<IConnectionPointContainer> cp_cont;

        auto hr = self->disp->ptr->QueryInterface(&cp_cont);
        if (SUCCEEDED(hr)) {
            CComPtr<IConnectionPoint> point;
            hr = cp_cont->FindConnectionPoint(it->second.clsid, &point);

            if (SUCCEEDED(hr) && point) {
                hr = point->Unadvise(dwCookie);
                if (SUCCEEDED(hr)) {
                    self->connections_.erase(it);
                }
                point.Release();
            }
            cp_cont.Release();			
        }
	}
}
