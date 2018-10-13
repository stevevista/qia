//-------------------------------------------------------------------------------------------------------
// Project: node-activex
// Author: Yuri Dursin
// Description: Defines the entry point for the NodeJS addon
//-------------------------------------------------------------------------------------------------------

#include "dispatch_object.h"
#include "dispatch_callback.h"


NAN_MODULE_INIT(init) {
    Nan::HandleScope scope;

    DispObject::NodeInit(target);
	VariantObject::NodeInit(target);
    DispatchCallback::Initialize(target);
}

NODE_MODULE(ole_bindings, init)

//----------------------------------------------------------------------------------

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ulReason, LPVOID lpReserved) {
	
    switch (ulReason) {
    case DLL_PROCESS_ATTACH:
        //CoInitialize(0);
        CoInitializeEx(0, COINIT_MULTITHREADED);
        break;
    case DLL_PROCESS_DETACH:
        CoUninitialize();
        break;
    case DLL_THREAD_ATTACH:
        break;
    case DLL_THREAD_DETACH:
        break;
    }
    return TRUE;
}

//----------------------------------------------------------------------------------
