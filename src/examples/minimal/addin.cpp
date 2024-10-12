#include <xll/xll.hpp>

XLL_EXPORT const char * __stdcall testFunction(xll::variant *)
{
    const char *result = "Success!";
    return result;
}

/// This function is called by the Add-in Manager to find the long name of the
/// add-in. If xAction = 1, this function should return a string containing the
/// long name of this XLL. If xAction = 2 or 3, this function should return
/// xlerrValue (#VALUE!).
XLL_EXPORT xll::variant * __stdcall xlAddInManagerInfo12(xll::variant *xAction)
{
    xll::log()->info("xlAddInManagerInfo12");

    using namespace xll;

    static variant xInfo, xIntAction;
    switch (xAction->xltype()) {
    case xltypeNum:
        xIntAction = static_cast<int>(xAction->get<xlnum>());
        break;
    case xltypeInt:
        xIntAction = xAction->get<xlint>();
        break;
    default:
        break;
    }
    
    if (xIntAction.xltype() == xltypeInt && xIntAction.get<xlint>() == 1)
        xInfo = L"Sample XLL";
    else
        xInfo = error::xlerrValue;
    
    return &xInfo;
}

/// Excel calls xlAutoOpen when the XLL is loaded. Register all functions and
/// perform any additional initialization in this function.
/// \return 1 on success, 0 on failure.
XLL_EXPORT int __stdcall xlAutoOpen()
{
    xll::log()->info("xlAutoOpen");

    auto opts = xll::function_options();
    opts.argument_text = L"arg";
    opts.category = L"Sample";
    opts.function_help = L"Sample function returning a string.";
    opts.argument_help = { L"Argument ignored." };
    xll::register_function(testFunction, L"testFunction", L"TEST.FUNCTION", opts);

    return 1;
}

/// Excel calls xlAutoClose when it unloads the XLL.
XLL_EXPORT int __stdcall xlAutoClose()
{
    xll::log()->info("xlAutoClose");
    return 1;
}

/// Called when the XLL is added to the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoOpen.
XLL_EXPORT int __stdcall xlAutoAdd()
{
    xll::log()->info("xlAutoAdd");
    return 1;
}

/// Called when the XLL is removed from the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoRemove() and xlfUnregister.
XLL_EXPORT int __stdcall xlAutoRemove()
{
    xll::log()->info("xlAutoRemove");
    return 1;
}

/// Free internally allocated arrays and call destructor.
XLL_EXPORT int __stdcall xlAutoFree12(xll::variant *pxFree)
{
    xll::log()->info("xlAutoFree12");
    if (pxFree)
        pxFree->release();
    return 1;
}