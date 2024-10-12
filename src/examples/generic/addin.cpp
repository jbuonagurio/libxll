#include <xll/xll.hpp>

XLL_EXPORT const char * __stdcall test_string(xll::variant *)
{
    const char *result = "Success!";
    return result;
}

XLL_EXPORT int test_dialog()
{
    using namespace xll;

    xlmulti dialog({
        {xlnil(), xlnil(), xlnil(), 372, 200, L"Sample Dialog", xlnil()},
        {1, 100, 170, 90, xlnil(), L"OK", xlnil()},
        {10, 150, 50, xlnil(), xlnil(), xlnil(), xlnil()}
    });

    xloper<xlmulti> *args = new xloper<xlmulti>(dialog);
    args->set_flags(xlbitDLLFree);

    variant result;
    Excel12(xlfDialogBox, &result, args);
    
    // Clean up by passing in the return value of dialog creation call.
    if (result.xltype() == xltypeBool && (bool)result.get<xlbool>() == false) {
        Excel12(xlfDialogBox, nullptr, &result);
    }

    Excel12(xlFree, nullptr, &result);
    return 1;
}

XLL_EXPORT int __stdcall get_stack_size(xll::variant *)
{
    return xll::stack_size();    
}

/// This function is called by the Add-in Manager to find the long name of the
/// add-in. If xAction = 1, this function should return a string containing the
/// long name of this XLL. If xAction = 2 or 3, this function should return
/// xlerrValue (#VALUE!).
XLL_EXPORT xll::variant * __stdcall xlAddInManagerInfo12(xll::variant *xAction)
{
    using namespace xll;
    
    thread_local variant xInfo, xIntAction;
    xloper<xlint> xDestType(static_cast<int>(xltypeInt));

    Excel12(xlCoerce, &xIntAction, xAction, &xDestType);
    
    if (xIntAction.xltype() == xltypeInt && xIntAction.get<xlint>() == 1) {
        xInfo.emplace<xlstr>(L"Generic");
    }
    else {
        xInfo.emplace<xlerr>(error::xlerrValue);
    }
    
    return &xInfo;
}

/// Excel calls xlAutoOpen when the XLL is loaded. Register all functions and
/// perform any additional initialization in this function.
/// \return 1 on success, 0 on failure.
XLL_EXPORT int __stdcall xlAutoOpen()
{
    using namespace xll;
    
    {
        auto opts = function_options();
        opts.argument_text = L"x";
        opts.category = L"Generic";
        opts.function_help = L"Sample function.";
        attribute_set<tag::thread_safe> attrs;
        register_function(test_string, L"test_string", L"TEST.STRING", opts, attrs);
    }

    {
        auto opts = function_options();
        opts.type = macro_type::command;
        register_function(test_dialog, L"test_dialog", L"TEST.DIALOG", opts);
    }

    {
        auto opts = function_options();
        opts.category = L"Generic";
        register_function(get_stack_size, L"get_stack_size", L"STACK.SIZE", opts);
    }

    return 1;
}

/// Excel calls xlAutoClose when it unloads the XLL.
XLL_EXPORT int __stdcall xlAutoClose()
{
    return 1;
}

/// Called when the XLL is added to the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoOpen.
XLL_EXPORT int __stdcall xlAutoAdd()
{
    return 1;
}

/// Called when the XLL is removed from the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoRemove() and xlfUnregister.
XLL_EXPORT int __stdcall xlAutoRemove()
{
    return 1;
}

/// Free internally allocated arrays and call destructor.
XLL_EXPORT int __stdcall xlAutoFree12(xll::variant *pxFree)
{
    if (pxFree)
        pxFree->release();
    return 1;
}
