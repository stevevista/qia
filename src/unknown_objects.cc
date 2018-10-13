#include "unknown_objects.h"
#include "utils.h"

const GUID CLSID_DispObjectImpl = { 0x9dce8520, 0x2efe, 0x48c0,{ 0xa0, 0xdc, 0x95, 0x1b, 0x29, 0x18, 0x72, 0xc0 } };


//-------------------------------------------------------------------------------------------------------
// DispArrayImpl implemetation

HRESULT STDMETHODCALLTYPE DispArrayImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	LPOLESTR name = rgszNames[0];
	if (wcscmp(name, L"length") == 0) *rgDispId = 1;
	else return DISP_E_UNKNOWNNAME;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispArrayImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	HRESULT hrcode = S_OK;
	UINT argcnt = pDispParams->cArgs;
	VARIANT *args = pDispParams->rgvarg;

	if ((var.vt & VT_ARRAY) == 0) return E_NOTIMPL;
	SAFEARRAY *arr = ((var.vt & VT_BYREF) != 0) ? *var.pparray : var.parray;

	switch (dispIdMember) {
	case 1: {
		if (pVarResult) {
			pVarResult->vt = VT_INT;
			pVarResult->intVal = (INT)(arr ? arr->rgsabound[0].cElements : 0);
		}
		return hrcode; }
	}
	return E_NOTIMPL;
}


//-------------------------------------------------------------------------------------------------------
// DispEnumImpl implemetation

HRESULT STDMETHODCALLTYPE DispEnumImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
    if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
    LPOLESTR name = rgszNames[0];
    if (wcscmp(name, L"Next") == 0) *rgDispId = 1;
    else if (wcscmp(name, L"Skip") == 0) *rgDispId = 2;
    else if (wcscmp(name, L"Reset") == 0) *rgDispId = 3;
    else if (wcscmp(name, L"Clone") == 0) *rgDispId = 4;
    else return DISP_E_UNKNOWNNAME;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE DispEnumImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
    HRESULT hrcode = S_OK;
    UINT argcnt = pDispParams->cArgs;
    VARIANT *args = pDispParams->rgvarg;
    switch (dispIdMember) {
    case 1: {
        CComVariant arr;
        ULONG fetched, celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        if (!pVarResult || celt == 0) hrcode = E_INVALIDARG;
        if SUCCEEDED(hrcode) hrcode = arr.ArrayCreate(VT_VARIANT, celt);
        if SUCCEEDED(hrcode) hrcode = ptr->Next(celt, arr.ArrayGet<VARIANT>(0), &fetched);
        if SUCCEEDED(hrcode) {
            if (fetched == 0) pVarResult->vt = VT_EMPTY;
            else if (fetched == 1) {
                VARIANT *v = arr.ArrayGet<VARIANT>(0);
                *pVarResult = *v;
                v->vt = VT_EMPTY;
            }
            else {
                if (fetched < celt) hrcode = arr.ArrayResize(fetched);
                if SUCCEEDED(hrcode) arr.Detach(pVarResult);
            }
        }
        return hrcode; }
    case 2: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        ULONG celt = (argcnt > 0) ? Variant2Int(args[argcnt - 1], (ULONG)1) : 1;
        return ptr->Skip(celt); 
        }
    case 3: {
        if (pVarResult) pVarResult->vt = VT_EMPTY;
        return ptr->Reset(); 
        }
    case 4: {
        if (!pVarResult) hrcode = E_INVALIDARG;
        std::auto_ptr<DispEnumImpl> disp;
        if SUCCEEDED(hrcode) hrcode = ptr->Clone(&disp->ptr);
        if SUCCEEDED(hrcode) {
            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = disp.release();
            pVarResult->pdispVal->AddRef();
        }
        return hrcode; }
    }
    return E_NOTIMPL;
}

//-------------------------------------------------------------------------------------------------------
// DispObjectImpl implemetation

HRESULT STDMETHODCALLTYPE DispObjectImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) {
	if (cNames != 1 || !rgszNames[0]) return DISP_E_UNKNOWNNAME;
	std::wstring name(rgszNames[0]);
	name_ptr &ptr = names[name];
	if (!ptr) {
		ptr.reset(new name_t(dispid_next++, name));
		index.insert(index_t::value_type(ptr->dispid, ptr));
	}
	*rgDispId = ptr->dispid;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE DispObjectImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) {
	Isolate *isolate = Isolate::GetCurrent();
	Local<Object> self = obj.Get(isolate);
	Local<Value> name, val, ret;

	// Prepare name by member id
	if (dispIdMember != DISPID_VALUE) {
		index_t::const_iterator p = index.find(dispIdMember);
		if (p == index.end()) return DISP_E_MEMBERNOTFOUND;
		name_t &info = *p->second;
		name = String::NewFromTwoByte(isolate, (uint16_t*)info.name.c_str());
	}

	// Set property value
	if ((wFlags & DISPATCH_PROPERTYPUT) != 0) {
		UINT argcnt = pDispParams->cArgs;
		VARIANT *key = (argcnt > 1) ? &pDispParams->rgvarg[--argcnt] : nullptr;
		if (argcnt > 0) val = Variant2Value(isolate, pDispParams->rgvarg[--argcnt], true);
		else val = Undefined(isolate);
		bool rcode;

		// Set simple object property value
		if (!key) {
			if (name.IsEmpty()) return DISP_E_MEMBERNOTFOUND;
			rcode = self->Set(name, val);
		}

		// Set object/array item value
		else {
			Local<Object> target;
			if (name.IsEmpty()) target = self;
			else {
				Local<Value> obj = self->Get(name);
				if (!obj.IsEmpty()) target = Local<Object>::Cast(obj);
				if (target.IsEmpty()) return DISP_E_BADCALLEE;
			}

			LONG index = Variant2Int<LONG>(*key, -1);
			if (index >= 0) rcode = target->Set((uint32_t)index, val);
			else rcode = target->Set(Variant2Value(isolate, *key, false), val);
		}

		// Store result
		if (pVarResult) {
			pVarResult->vt = VT_BOOL;
			pVarResult->boolVal = rcode ? VARIANT_TRUE : VARIANT_FALSE;
		}
		return S_OK;
	}

	// Prepare property item
	if (name.IsEmpty()) val = self;
	else val = self->Get(name);

	// Call property as method
	if ((wFlags & DISPATCH_METHOD) != 0) {
		wFlags = 0;
		NodeArguments args(isolate, pDispParams, true);
		int argcnt = (int)args.items.size();
		Local<Value> *argptr = (argcnt > 0) ? &args.items[0] : nullptr;
		if (val->IsFunction()) {
			Local<Function> func = Local<Function>::Cast(val);
			if (func.IsEmpty()) return DISP_E_BADCALLEE;
			ret = func->Call(self, argcnt, argptr);
		}
		else if (val->IsObject()) {
			wFlags = DISPATCH_PROPERTYGET;
			//Local<Object> target = val->ToObject();
			//target->CallAsFunction(isolate->GetCurrentContext(), target, args.items.size(), &args.items[0]).ToLocal(&ret);
		}
		else {
			ret = val;
		}
	}

	// Get property value
	if ((wFlags & DISPATCH_PROPERTYGET) != 0) {
		if (pDispParams->cArgs == 1) {
			Local<Object> target;
			if (!val.IsEmpty()) target = Local<Object>::Cast(val);
			if (target.IsEmpty()) return DISP_E_BADCALLEE;
			VARIANT &key = pDispParams->rgvarg[0];
			LONG index = Variant2Int<LONG>(key, -1);
			if (index >= 0) ret = target->Get((uint32_t)index);
			else ret = target->Get(Variant2Value(isolate, key, false));
		}
		else {
			ret = val;
		}
	}

	// Store result
	if (pVarResult) {
		Value2Variant(isolate, ret, *pVarResult, VT_NULL);
	}
	return S_OK;
}
