//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#pragma pack(push, 8)

#include <comdef.h>

// Generated from Microsoft Office type libraries:
// 
// Microsoft Add-In Designer (Ver 1.0)
// GUID: {AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}
// Path: "%PROGRAMFILES(X86)%\Common Files\Designer\MSADDNDR.OLB"
// #import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" raw_interfaces_only raw_native_types named_guids
//
// Microsoft Office 16.0 Object Library
// GUID: {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}
// Path: "%PROGRAMFILES(X86)%\Microsoft Office\Root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\mso.dll"
// #import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" raw_interfaces_only raw_native_types named_guids
//
//
// Microsoft Excel 16.0 Object Library (Ver 1.9)
// GUID: {00020813-0000-0000-C000-000000000046}
// Path: "%PROGRAMFILES(X86)%\Microsoft Office\root\Office16\EXCEL.EXE"
// #import "libid:00020813-0000-0000-C000-000000000046" raw_interfaces_only raw_native_types named_guids auto_search rename("RGB", "MsoRGB") rename("DialogBox", "MsoDialogBox")

//
// Forward references and typedefs
//

struct __declspec(uuid("ac0714f2-3d04-11d1-ae7d-00a0c90f26f4")) /* LIBID */ __AddInDesignerObjects;
struct __declspec(uuid("2df8d04c-5bfa-101b-bde5-00aa0044de52")) /* LIBID */ __Office;

struct __declspec(uuid("b65ad801-abaf-11d0-bb8b-00a0c90f2744")) /* dual interface */ _IDTExtensibility2;
struct __declspec(uuid("000c033e-0000-0000-c000-000000000046")) /* dual interface */ ICustomTaskPaneConsumer;
struct __declspec(uuid("000c03a7-0000-0000-c000-000000000046")) /* dual interface */ IRibbonUI;
struct __declspec(uuid("000c0395-0000-0000-c000-000000000046")) /* dual interface */ IRibbonControl;
struct __declspec(uuid("000c0396-0000-0000-c000-000000000046")) /* dual interface */ IRibbonExtensibility;

typedef struct _IDTExtensibility2 * IDTExtensibility2;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(_IDTExtensibility2, __uuidof(_IDTExtensibility2));
_COM_SMARTPTR_TYPEDEF(ICustomTaskPaneConsumer, __uuidof(ICustomTaskPaneConsumer));
_COM_SMARTPTR_TYPEDEF(IRibbonUI, __uuidof(IRibbonUI));
_COM_SMARTPTR_TYPEDEF(IRibbonControl, __uuidof(IRibbonControl));
_COM_SMARTPTR_TYPEDEF(IRibbonExtensibility, __uuidof(IRibbonExtensibility));

//
// Type library items
//

enum ext_ConnectMode {
    ext_cm_AfterStartup = 0,
    ext_cm_Startup = 1,
    ext_cm_External = 2,
    ext_cm_CommandLine = 3
};

enum ext_DisconnectMode {
    ext_dm_HostShutdown = 0,
    ext_dm_UserClosed = 1
};

struct __declspec(uuid("b65ad801-abaf-11d0-bb8b-00a0c90f2744")) _IDTExtensibility2 : IDispatch
{
    virtual HRESULT __stdcall OnConnection(
      /*[in]*/ IDispatch *Application,
      /*[in]*/ ext_ConnectMode ConnectMode,
      /*[in]*/ IDispatch *AddInInst,
      /*[in]*/ SAFEARRAY **custom) = 0;
    virtual HRESULT __stdcall OnDisconnection(
      /*[in]*/ ext_DisconnectMode RemoveMode,
      /*[in]*/ SAFEARRAY **custom) = 0;
    virtual HRESULT __stdcall OnAddInsUpdate(
      /*[in]*/ SAFEARRAY **custom) = 0;
    virtual HRESULT __stdcall OnStartupComplete(
      /*[in]*/ SAFEARRAY **custom) = 0;
    virtual HRESULT __stdcall OnBeginShutdown(
      /*[in]*/ SAFEARRAY **custom) = 0;
};

struct __declspec(uuid("000c0396-0000-0000-c000-000000000046")) IRibbonExtensibility : IDispatch
{
    virtual HRESULT __stdcall GetCustomUI(
      /*[in]*/ BSTR RibbonID,
      /*[out,retval]*/ BSTR *RibbonXml) = 0;
};

//
// Named GUID constants initializations
//

static constexpr GUID LIBID_AddInDesignerObjects =
    {0xac0714f2,0x3d04,0x11d1,{0xae,0x7d,0x00,0xa0,0xc9,0x0f,0x26,0xf4}};
static constexpr GUID LIBID_Office =
    {0x2df8d04c,0x5bfa,0x101b,{0xbd,0xe5,0x00,0xaa,0x00,0x44,0xde,0x52}};

static constexpr GUID IID__IDTExtensibility2 =
    {0xb65ad801,0xabaf,0x11d0,{0xbb,0x8b,0x00,0xa0,0xc9,0x0f,0x27,0x44}};
static constexpr GUID IID_ICustomTaskPaneConsumer =
    {0x000c033e,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IRibbonUI =
    {0x000c03a7,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IRibbonControl =
    {0x000c0395,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IRibbonExtensibility =
    {0x000c0396,0x0000,0x0000,{0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

#pragma pack(pop)
