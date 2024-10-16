//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

/**
 * \file log.hpp
 * OutputDebugStringA logging using spdlog. Useful with Sysinternals DebugView.
 */

#include <xll/config.hpp>

#ifndef XLL_USE_SPDLOG

#include <memory>
#include <string_view>

namespace xll {
namespace detail {

struct log_impl {
    template<class... Args> static void trace(std::string_view, const Args&...) {}
    template<class... Args> static void debug(std::string_view, const Args&...) {}
    template<class... Args> static void info(std::string_view, const Args&...) {}
    template<class... Args> static void warn(std::string_view, const Args&...) {}
    template<class... Args> static void error(std::string_view, const Args&...) {}
    template<class... Args> static void critical(std::string_view, const Args&...) {}
};

} // namespace detail

inline auto log() {
    static auto logger = std::make_shared<detail::log_impl>();
    return logger;
}

} // namespace xll

#else

#include <xll/constants.hpp>

#if BOOST_OS_WINDOWS
#include <boost/winapi/debugapi.hpp>
#include <boost/winapi/get_current_process_id.hpp>
#include <boost/winapi/get_current_thread_id.hpp>
#else
#include <sys/types.h>
#include <unistd.h>
#include <pthread.h>
#include <iostream>
#endif

#include <spdlog/spdlog.h>

#if BOOST_OS_WINDOWS
#include <spdlog/sinks/base_sink.h>
#else
#include <spdlog/sinks/syslog_sink.h>
#endif

#include <memory>
#include <mutex>

namespace fmt {

template<>
struct formatter<xll::XLTYPE>
{
    template<typename ParseContext>
    constexpr auto parse(ParseContext &ctx) { return ctx.begin(); }

    template<typename FormatContext>
    auto format(xll::XLTYPE xltype, FormatContext &ctx)
    {
        switch (xltype) {
        case xll::XLTYPE::xltypeNum:                   return format_to(ctx.out(), "xltypeNum");
        case xll::XLTYPE::xltypeStr:                   return format_to(ctx.out(), "xltypeStr");
        case xll::XLTYPE::xltypeBool:                  return format_to(ctx.out(), "xltypeBool");
        case xll::XLTYPE::xltypeRef:                   return format_to(ctx.out(), "xltypeRef");
        case xll::XLTYPE::xltypeErr:                   return format_to(ctx.out(), "xltypeErr");
        case xll::XLTYPE::xltypeFlow:                  return format_to(ctx.out(), "xltypeFlow");
        case xll::XLTYPE::xltypeMulti:                 return format_to(ctx.out(), "xltypeMulti");
        case xll::XLTYPE::xltypeMissing:               return format_to(ctx.out(), "xltypeMissing");
        case xll::XLTYPE::xltypeNil:                   return format_to(ctx.out(), "xltypeNil");
        case xll::XLTYPE::xltypeSRef:                  return format_to(ctx.out(), "xltypeSRef");
        case xll::XLTYPE::xltypeInt:                   return format_to(ctx.out(), "xltypeInt");
        case xll::XLTYPE::xltypeBigData:               return format_to(ctx.out(), "xltypeBigData");
        default:                                       return format_to(ctx.out(), static_cast<int>(xltype));
        }
    }
};

template<>
struct formatter<xll::XLRET>
{
    template<typename ParseContext>
    constexpr auto parse(ParseContext &ctx) { return ctx.begin(); }

    template<typename FormatContext>
    auto format(xll::XLRET rc, FormatContext &ctx)
    {
        switch (rc) {
        case xll::XLRET::xlretSuccess:                 return format_to(ctx.out(), "xlretSuccess");
        case xll::XLRET::xlretAbort:                   return format_to(ctx.out(), "xlretAbort");
        case xll::XLRET::xlretInvXlfn:                 return format_to(ctx.out(), "xlretInvXlfn");
        case xll::XLRET::xlretInvCount:                return format_to(ctx.out(), "xlretInvCount");
        case xll::XLRET::xlretInvXloper:               return format_to(ctx.out(), "xlretInvXloper");
        case xll::XLRET::xlretStackOvfl:               return format_to(ctx.out(), "xlretStackOvfl");
        case xll::XLRET::xlretFailed:                  return format_to(ctx.out(), "xlretFailed");
        case xll::XLRET::xlretUncalced:                return format_to(ctx.out(), "xlretUncalced");
        case xll::XLRET::xlretNotThreadSafe:           return format_to(ctx.out(), "xlretNotThreadSafe");
        case xll::XLRET::xlretInvAsynchronousContext:  return format_to(ctx.out(), "xlretInvAsynchronousContext");
        case xll::XLRET::xlretNotClusterSafe:          return format_to(ctx.out(), "xlretNotClusterSafe");
        default:                                       return format_to(ctx.out(), static_cast<int>(rc));
        }
    }
};

} // namespace fmt

namespace xll {
namespace sinks {

#if BOOST_OS_WINDOWS

template<typename Mutex>
class debug_sink : public spdlog::sinks::base_sink<Mutex>
{
public:
    explicit debug_sink() {}
protected:
    void sink_it_(const spdlog::details::log_msg& msg) override
    {
        spdlog::memory_buf_t formatted;
        base_sink<Mutex>::formatter_->format(msg, formatted);
        boost::winapi::OutputDebugStringA(fmt::to_string(formatted).c_str());
    }
    void flush_() override {}
};

using debug_sink_mt = debug_sink<std::mutex>;

#endif

} // namespace sinks

inline int process_id()
{
#if BOOST_OS_WINDOWS
    return static_cast<int>(boost::winapi::GetCurrentProcessId());
#else
    return static_cast<int>(::getpid());
#endif
}

inline std::size_t thread_id()
{
#if BOOST_OS_WINDOWS
    return static_cast<std::size_t>(boost::winapi::GetCurrentThreadId());
#else
    uint64_t tid;
    pthread_threadid_np(nullptr, &tid);
    return static_cast<std::size_t>(tid);
#endif
}

inline auto log()
{
#if BOOST_OS_WINDOWS
    static auto sink = std::make_shared<sinks::debug_sink_mt>();
    static auto logger = std::make_shared<spdlog::logger>("xll", sink);
#else    
    static auto logger = spdlog::syslog_logger_mt("xll");
#endif

    std::once_flag flag;
    std::call_once(flag, [](){
        logger->set_pattern("[%t] [%T.%e] [%l] %v"); // thread, time, level
        logger->set_level(spdlog::level::debug);
    });

    return logger;
}

} // namespace xll

#endif // XLL_USE_SPDLOG