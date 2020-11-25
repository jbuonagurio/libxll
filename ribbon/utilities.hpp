//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <interfaces.hpp>

#include <string>
#include <cstdio>

inline std::string GuidToStdString(const GUID& guid)
{
    char buffer[39];
    const char *format = "{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}";
    snprintf(buffer, sizeof(buffer), format,
        guid.Data1, guid.Data2, guid.Data3,
        guid.Data4[0], guid.Data4[1], guid.Data4[2],
        guid.Data4[3], guid.Data4[4], guid.Data4[5],
        guid.Data4[6], guid.Data4[7]);
    return std::string(buffer);
}

inline std::string GetIIDString(const IID& iid)
{
    if (iid == __uuidof(IUnknown))
        return "IID_IUnknown";
    else if (iid == __uuidof(IDispatch))
        return "IID_IDispatch";
    else if (iid == __uuidof(_IDTExtensibility2))
        return "IID__IDTExtensibility2";
    else if (iid == __uuidof(IRibbonExtensibility))
        return "IID_IRibbonExtensibility";
    else if (iid == __uuidof(IRibbonCallback))
        return "IID_IRibbonCallback";
    else if (iid == __uuidof(ICustomTaskPaneConsumer))
        return "IID_ICustomTaskPaneConsumer";
    else if (iid == IID_NULL)
        return "IID_NULL";
    else
        return GuidToStdString(iid);
}
