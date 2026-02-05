#include "DesktopManagerWrapper.h"

Napi::FunctionReference* DesktopManagerWrapper::Initialize(Napi::Env env)
{
    Napi::Function func = DefineClass(env, "DesktopManager", {
        InstanceMethod("createDesktopInWindow", &DesktopManagerWrapper::CreateDesktopInWindow),
        InstanceMethod("getChildSession", &DesktopManagerWrapper::GetChildSession),
        InstanceMethod("getRDPClient", &DesktopManagerWrapper::GetRDPClient),
        InstanceMethod("launchApplication", &DesktopManagerWrapper::LaunchApplication),
        InstanceMethod("resizeDesktop", &DesktopManagerWrapper::ResizeDesktop),
        InstanceMethod("cleanup", &DesktopManagerWrapper::Cleanup),
        InstanceMethod("getSessionId", &DesktopManagerWrapper::GetSessionId),
        InstanceMethod("isConnected", &DesktopManagerWrapper::IsConnected),
    });

    Napi::FunctionReference* constructor = new Napi::FunctionReference();
    *constructor = Napi::Persistent(func);

    return constructor;
}

DesktopManagerWrapper::DesktopManagerWrapper(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<DesktopManagerWrapper>(info), sessionId(0), isManaged(false)
{
}

DesktopManagerWrapper::~DesktopManagerWrapper()
{
    if (isManaged)
    {
        Cleanup(Napi::CallbackInfo(Env(), nullptr));
    }
}

Napi::Value DesktopManagerWrapper::CreateDesktopInWindow(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1)
    {
        throw Napi::TypeError::New(env, "Window handle expected");
    }

    auto addon = env.GetInstanceData<AutomationAddon>();
    
    Napi::Object childSessionObj = addon->ChildSessionWrapperConstructor->New({});
    ChildSessionWrapper* childSession = ChildSessionWrapper::Unwrap(childSessionObj);
    
    Napi::Value sessionIdValue = childSession->GetSessionId(info);
    sessionId = sessionIdValue.As<Napi::Number>().Uint32Value();
    
    childSessionRef = Napi::Reference<Napi::Object>::New(childSessionObj, 1);

    Napi::Object rdpClientObj = addon->RDPClientWrapperConstructor->New({});
    RDPClientWrapper* rdpClient = RDPClientWrapper::Unwrap(rdpClientObj);
    
    rdpClientRef = Napi::Reference<Napi::Object>::New(rdpClientObj, 1);

    rdpClient->SetServer(Napi::CallbackInfo(env, info.This(), 1, 
        new Napi::Value[1]{Napi::String::New(env, "localhost")}, nullptr));
    
    rdpClient->SetPort(Napi::CallbackInfo(env, info.This(), 1, 
        new Napi::Value[1]{Napi::Number::New(env, 3389)}, nullptr));

    if (info.Length() >= 3 && info[1].IsNumber() && info[2].IsNumber())
    {
        long width = info[1].As<Napi::Number>().Int32Value();
        long height = info[2].As<Napi::Number>().Int32Value();
        
        Napi::Value args[2] = {Napi::Number::New(env, width), Napi::Number::New(env, height)};
        rdpClient->SetDesktopSize(Napi::CallbackInfo(env, info.This(), 2, args, nullptr));
    }
    else
    {
        Napi::Value args[2] = {Napi::Number::New(env, 1024), Napi::Number::New(env, 768)};
        rdpClient->SetDesktopSize(Napi::CallbackInfo(env, info.This(), 2, args, nullptr));
    }

    rdpClient->SetFullscreen(Napi::CallbackInfo(env, info.This(), 1, 
        new Napi::Value[1]{Napi::Boolean::New(env, false)}, nullptr));
    
    rdpClient->SetColorDepth(Napi::CallbackInfo(env, info.This(), 1, 
        new Napi::Value[1]{Napi::Number::New(env, 32)}, nullptr));
    
    rdpClient->SetAuthenticationLevel(Napi::CallbackInfo(env, info.This(), 1, 
        new Napi::Value[1]{Napi::Number::New(env, 0)}, nullptr));

    rdpClient->EmbedInWindow(info);

    Sleep(500);

    rdpClient->Connect(Napi::CallbackInfo(env, info.This(), 0, nullptr, nullptr));

    isManaged = true;

    Napi::Object result = Napi::Object::New(env);
    result.Set("sessionId", Napi::Number::New(env, sessionId));
    result.Set("childSession", childSessionObj);
    result.Set("rdpClient", rdpClientObj);

    return result;
}

Napi::Value DesktopManagerWrapper::GetChildSession(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (childSessionRef.IsEmpty())
    {
        return env.Null();
    }

    return childSessionRef.Value();
}

Napi::Value DesktopManagerWrapper::GetRDPClient(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (rdpClientRef.IsEmpty())
    {
        return env.Null();
    }

    return rdpClientRef.Value();
}

Napi::Value DesktopManagerWrapper::LaunchApplication(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (childSessionRef.IsEmpty())
    {
        throw Napi::Error::New(env, "Desktop not initialized");
    }

    Napi::Object childSessionObj = childSessionRef.Value();
    ChildSessionWrapper* childSession = ChildSessionWrapper::Unwrap(childSessionObj);

    return childSession->LaunchProcess(info);
}

Napi::Value DesktopManagerWrapper::ResizeDesktop(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (rdpClientRef.IsEmpty())
    {
        throw Napi::Error::New(env, "RDP client not initialized");
    }

    if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsNumber())
    {
        throw Napi::TypeError::New(env, "Two numbers expected (width, height)");
    }

    Napi::Object rdpClientObj = rdpClientRef.Value();
    RDPClientWrapper* rdpClient = RDPClientWrapper::Unwrap(rdpClientObj);

    return rdpClient->SetDesktopSize(info);
}

Napi::Value DesktopManagerWrapper::Cleanup(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (!rdpClientRef.IsEmpty())
    {
        Napi::Object rdpClientObj = rdpClientRef.Value();
        RDPClientWrapper* rdpClient = RDPClientWrapper::Unwrap(rdpClientObj);
        rdpClient->Disconnect(Napi::CallbackInfo(env, info.This(), 0, nullptr, nullptr));
        rdpClientRef.Reset();
    }

    if (!childSessionRef.IsEmpty())
    {
        Napi::Object childSessionObj = childSessionRef.Value();
        ChildSessionWrapper* childSession = ChildSessionWrapper::Unwrap(childSessionObj);
        childSession->Terminate(Napi::CallbackInfo(env, info.This(), 0, nullptr, nullptr));
        childSessionRef.Reset();
    }

    isManaged = false;

    return Napi::Boolean::New(env, true);
}

Napi::Value DesktopManagerWrapper::GetSessionId(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    return Napi::Number::New(env, sessionId);
}

Napi::Value DesktopManagerWrapper::IsConnected(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();

    if (rdpClientRef.IsEmpty())
    {
        return Napi::Boolean::New(env, false);
    }

    Napi::Object rdpClientObj = rdpClientRef.Value();
    RDPClientWrapper* rdpClient = RDPClientWrapper::Unwrap(rdpClientObj);

    Napi::Value state = rdpClient->GetConnectionState(info);
    return Napi::Boolean::New(env, state.As<Napi::Number>().Int32Value() > 0);
}
