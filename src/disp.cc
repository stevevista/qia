//-------------------------------------------------------------------------------------------------------
// Project: NodeActiveX
// Author: Yuri Dursin
// Description: DispObject class implementations
//-------------------------------------------------------------------------------------------------------

#include "disp.h"

Persistent<ObjectTemplate> VariantObject::inst_template;
Persistent<FunctionTemplate> VariantObject::clazz_template;


//-------------------------------------------------------------------------------------------------------

class vtypes_t {
public:
    inline vtypes_t(std::initializer_list<std::pair<std::wstring, VARTYPE>> recs) {
        for (auto &rec : recs) {
            str2vt.emplace(rec.first, rec.second);
            vt2str.emplace(rec.second, rec.first);
        }
    }
    inline bool find(VARTYPE vt, std::wstring &name) {
        auto it = vt2str.find(vt);
        if (it == vt2str.end()) return false;
        name = it->second;
        return true;
    }
    inline VARTYPE find(const std::wstring &name) {
        auto it = str2vt.find(name);
        if (it == str2vt.end()) return VT_EMPTY;
        return it->second;
    }
private:
    std::map<std::wstring, VARTYPE> str2vt;
    std::map<VARTYPE, std::wstring> vt2str;
};

static vtypes_t vtypes({
	{ L"char", VT_I1 },
	{ L"uchar", VT_UI1 },
	{ L"byte", VT_UI1 },
	{ L"short", VT_I2 },
	{ L"ushort", VT_UI2 },
	{ L"int", VT_INT },
	{ L"uint", VT_UINT },
	{ L"long", VT_I8 },
	{ L"ulong", VT_UI8 },

	{ L"int8", VT_I1 },
	{ L"uint8", VT_UI1 },
	{ L"int16", VT_I2 },
	{ L"uint16", VT_UI2 },
	{ L"int32", VT_I4 },
	{ L"uint32", VT_UI4 },
	{ L"int64", VT_I8 },
	{ L"uint64", VT_UI8 },
	{ L"currency", VT_CY },

	{ L"float", VT_R4 },
	{ L"double", VT_R8 },
	{ L"date", VT_DATE },
	{ L"decimal", VT_DECIMAL },

	{ L"string", VT_BSTR },
	{ L"empty", VT_EMPTY },
	{ L"variant", VT_VARIANT },
	{ L"null", VT_NULL },
	{ L"byref", VT_BYREF }
});

bool VariantObject::assign(Isolate *isolate, Local<Value> &val, Local<Value> &type) {
	VARTYPE vt = VT_EMPTY;
	if (!type.IsEmpty()) {
		if (type->IsString()) {
			Local<String> vtval = type->ToString();
			String::Value vtstr(vtval);
			const wchar_t *pvtstr = (const wchar_t *)*vtstr;
			int vtstr_len = vtstr.length();
			if (vtstr_len > 0 && *pvtstr == 'p') {
				vt |= VT_BYREF;
				vtstr_len--;
				pvtstr++;
			}
			if (vtstr_len > 0) {
				std::wstring type(pvtstr, vtstr_len);
                vt |= vtypes.find(type);
			}
		}
		else if (type->IsInt32()) {
			vt |= type->Int32Value();
		}
	}

	if (val.IsEmpty()) {
		if FAILED(value.ChangeType(vt)) return false;
		if ((value.vt & VT_BYREF) == 0) pvalue.Clear();
		return true;
	}

	value.Clear();
	pvalue.Clear();
	if ((vt & VT_BYREF) == 0) {
		Value2Variant(isolate, val, value, vt);
	}
	else {
		VARIANT *refvalue = nullptr;
		VARTYPE vt_noref = vt & ~VT_BYREF;
		VariantObject *ref = (!val.IsEmpty() && val->IsObject()) ? GetInstanceOf(isolate, val->ToObject()) : nullptr;
		if (ref) {
			if ((ref->value.vt & VT_BYREF) != 0) value = ref->value;
			else refvalue = &ref->value;
		}
		else {
			Value2Variant(isolate, val, pvalue, vt_noref);
			refvalue = &pvalue;
		}
		if (refvalue) {
			if (vt_noref == 0 || vt_noref == VT_VARIANT || refvalue->vt == VT_EMPTY) {
				value.vt = VT_VARIANT | VT_BYREF;
				value.pvarVal = refvalue;
			}
			else {
				value.vt = refvalue->vt | VT_BYREF;
				value.byref = &refvalue->intVal;
			}
		}
	}
	return true;
}

VariantObject::VariantObject(const FunctionCallbackInfo<Value> &args) {
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) val = args[0];
	if (argcnt > 1) type = args[1];
	assign(args.GetIsolate(), val, type);
}

void VariantObject::NodeInit(const Local<Object> &target) {
	Isolate *isolate = target->GetIsolate();

	// Prepare constructor template
	Local<FunctionTemplate> clazz = FunctionTemplate::New(isolate, NodeCreate);
	clazz->SetClassName(String::NewFromUtf8(isolate, "Variant"));

	NODE_SET_PROTOTYPE_METHOD(clazz, "clear", NodeClear);
	NODE_SET_PROTOTYPE_METHOD(clazz, "assign", NodeAssign);
	NODE_SET_PROTOTYPE_METHOD(clazz, "cast", NodeCast);
	NODE_SET_PROTOTYPE_METHOD(clazz, "toString", NodeToString);
	NODE_SET_PROTOTYPE_METHOD(clazz, "valueOf", NodeValueOf);

	Local<ObjectTemplate> &inst = clazz->InstanceTemplate();
	inst->SetInternalFieldCount(1);
	inst->SetNamedPropertyHandler(NodeGet, NodeSet);
	inst->SetIndexedPropertyHandler(NodeGetByIndex, NodeSetByIndex);
	//inst->SetCallAsFunctionHandler(NodeCall);
	//inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__id"), NodeGet);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__value"), NodeGet);
	inst->SetNativeDataProperty(String::NewFromUtf8(isolate, "__type"), NodeGet);

	inst_template.Reset(isolate, inst);
	clazz_template.Reset(isolate, clazz);
	target->Set(String::NewFromUtf8(isolate, "Variant"), clazz->GetFunction());
	NODE_DEBUG_MSG("VariantObject initialized");
}

Local<Object> VariantObject::NodeCreateInstance(const FunctionCallbackInfo<Value> &args) {
	Local<Object> self;
	Isolate *isolate = args.GetIsolate();
	if (!inst_template.IsEmpty()) {
		self = inst_template.Get(isolate)->NewInstance();
		(new VariantObject(args))->Wrap(self);
	}
	return self;
}

void VariantObject::NodeCreate(const FunctionCallbackInfo<Value> &args) {
	Isolate *isolate = args.GetIsolate();
	Local<Object> &self = args.This();
	(new VariantObject(args))->Wrap(self);
	args.GetReturnValue().Set(self);
}

void VariantObject::NodeClear(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	self->value.Clear();
	self->pvalue.Clear();
}

void VariantObject::NodeAssign(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) val = args[0];
	if (argcnt > 1) type = args[1];
	self->assign(isolate, val, type);
}

void VariantObject::NodeCast(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> val, type;
	int argcnt = args.Length();
	if (argcnt > 0) type = args[0];
	self->assign(isolate, val, type);
}

void VariantObject::NodeValueOf(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result = Variant2Value(isolate, self->value, true);
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeToString(const FunctionCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result = Variant2String(isolate, self->value);
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeGet(Local<String> name, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	String::Value vname(name);
	LPOLESTR id = (vname.length() > 0) ? (LPOLESTR)*vname : L"";
	if (_wcsicmp(id, L"__value") == 0) {
		Local<Value> result = Variant2Value(isolate, self->value);
		args.GetReturnValue().Set(result);
	}
	else if (_wcsicmp(id, L"__type") == 0) {
		std::wstring type, name;
		if (self->value.vt & VT_BYREF) type += L"byref:";
		if (self->value.vt & VT_ARRAY) type = L"array:";
        if (vtypes.find(self->value.vt & VT_TYPEMASK, name)) type += name;
		else type += std::to_wstring(self->value.vt & VT_TYPEMASK);
		args.GetReturnValue().Set(String::NewFromTwoByte(isolate, (uint16_t*)type.c_str()));
	}
	else if (_wcsicmp(id, L"__proto__") == 0) {
		Local<FunctionTemplate> clazz = clazz_template.Get(isolate);
		if (clazz.IsEmpty()) args.GetReturnValue().SetNull();
		else args.GetReturnValue().Set(clazz_template.Get(isolate)->GetFunction());
	}
	else if (_wcsicmp(id, L"clear") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeClear, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"assign") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeAssign, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"cast") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeCast, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"valueOf") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeValueOf, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"toString") == 0) {
		args.GetReturnValue().Set(FunctionTemplate::New(isolate, NodeToString, args.This())->GetFunction());
	}
	else if (_wcsicmp(id, L"length") == 0) {
		if ((self->value.vt & VT_ARRAY) != 0) {
			args.GetReturnValue().Set((uint32_t)self->value.ArrayLength());
		}
	}
}

void VariantObject::NodeGetByIndex(uint32_t index, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	Local<Value> result;
	if ((self->value.vt & VT_ARRAY) == 0) {
		 result = Variant2Value(isolate, self->value);
	}
	else {
		CComVariant value;
		if SUCCEEDED(self->value.ArrayGet((LONG)index, value)) {
			result = Variant2Value(isolate, value);
		}
	}
	args.GetReturnValue().Set(result);
}

void VariantObject::NodeSet(Local<String> name, Local<Value> val, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	isolate->ThrowException(DispError(isolate, E_NOTIMPL));
}

void VariantObject::NodeSetByIndex(uint32_t index, Local<Value> value, const PropertyCallbackInfo<Value>& args) {
	Isolate *isolate = args.GetIsolate();
	VariantObject *self = VariantObject::Unwrap<VariantObject>(args.This());
	if (!self) {
		isolate->ThrowException(DispErrorInvalid(isolate));
		return;
	}
	isolate->ThrowException(DispError(isolate, E_NOTIMPL));
}
