//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/attributes.hpp>
#include <xll/constants.hpp>
#include <xll/callback.hpp>
#include <xll/functions.hpp>
#include <xll/xloper12.hpp>
#include <xll/detail/type_text.hpp>
#include <xll/log.hpp>

#include <algorithm>
#include <array>
#include <functional>
#include <mutex>
#include <string>
#include <type_traits>
#include <typeindex>
#include <typeinfo>
#include <unordered_map>

namespace xll {

struct function_options
{
    std::wstring argument_text;
    int macro_type = 1;
    std::wstring category;
    std::wstring shortcut_text;
    std::wstring help_topic;
    std::wstring function_help;
    std::vector<std::wstring> argument_help;
};

struct registry
{
private:
    static inline std::mutex mutex_;
    static inline std::unordered_map<std::type_index, double> data_;

public:
    template <class F>
    static bool add(F ptr, double id)
    {
        std::lock_guard<std::mutex> lock(mutex_);
        auto index = std::type_index(typeid(ptr));
        data_[index] = id;
        return true;
    }

    template <class F>
    static bool remove(F ptr)
    {
        std::lock_guard<std::mutex> lock(mutex_);
        auto index = std::type_index(typeid(ptr));
        return data_.erase(index);
    }
};

// Index    Name                   Type        Alt. Type
// -----    -------------------    ---------   ---------
//     0    pxModuleText           xltypeStr   
//     1    pxProcedure            xltypeStr   xltypeNum
//     2    pxTypeText             xltypeStr   
//     3    pxFunctionText         xltypeStr   
//     4    pxArgumentText         xltypeStr   
//     5    pxMacroType            xltypeInt   xltypeNum
//     6    pxCategory             xltypeStr   xltypeNum
//     7    pxShortcutText         xltypeStr   
//     8    pxHelpTopic            xltypeStr   
//     9    pxFunctionHelp         xltypeStr   
//    10    pxArgumentHelp[245]    xltypeStr   

template <class F, class... Tags>
inline double register_function(F ptr, const std::wstring& dll_alias,
    const std::wstring& function_text, const function_options& opts = {},
    const attribute_set<Tags...>& attrs = {})
{
    std::array<variant12, 255> args;

    std::once_flag flag;
    std::call_once(flag, [&](){
        Excel12(xlGetName, &args[0]);
    });

    args[1].emplace<xlstr>(dll_alias);
    args[2].emplace<xlstr>(make_wpstring_array(detail::type_text(ptr, attrs)));
    args[3].emplace<xlstr>(function_text);
    args[4].emplace<xlstr>(opts.argument_text);
    args[5].emplace<xlint>(opts.macro_type);
    args[6].emplace<xlstr>(opts.category);
    args[7].emplace<xlstr>(opts.shortcut_text);
    args[8].emplace<xlstr>(opts.help_topic);
    
    std::size_t nargs = 9;
    
    if (!opts.function_help.empty()) {
        args[9].emplace<xlstr>(opts.function_help);
        nargs++;
        for (std::size_t i = 0; i < std::min((std::size_t)245, opts.argument_help.size()); ++i) {
            args[10 + i].emplace<xlstr>(opts.argument_help[i]);
            nargs++;
        }
    }

    std::array<variant12 *, 255> pargs;
    for (std::size_t i = 0; i < nargs; ++i) {
        pargs[i] = &args[i];
    }
    
    variant12 idvar;
    XLRET rc = Excel12v(xlfRegister, &idvar, pargs, nargs);
    if (rc != XLRET::xlretSuccess || idvar.xltype() != xltypeNum) {
        xll::log()->error("Registration failed: {}", rc);
        return 0.0;
    }

    double id = static_cast<double>(idvar.get<xlnum>());
    registry::add(ptr, id);
    return id;
}

template <class F>
inline bool unregister_function(F ptr)
{
    return registry::remove(ptr);
}

} // namespace xll