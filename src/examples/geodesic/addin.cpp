#define _USE_MATH_DEFINES

#include <xll/xll.hpp>

#include <cmath>

#include <GeographicLib/Config.h>
#include <GeographicLib/Geodesic.hpp>
#include <GeographicLib/Constants.hpp>

/// Solves the direct geodesic problem on the WGS84 ellipsoid.
/// \param[in] lon1 longitude at origin (degrees).
/// \param[in] lat1 latitude at origin (degrees).
/// \param[in] x2 x-coordinate at offset (meters).
/// \param[in] y2 y-coordinate at offset (meters).
/// \return array containing latitude and longitude at offset (degrees).
XLL_EXPORT xll::static_fp12<2> * __stdcall geodesicForward(
    const double lon1, const double lat1, const double x2, const double y2)
{
    thread_local xll::static_fp12<2> result(1, 2);
    
    double azi1 = std::atan2(x2, y2) * 180.0 / M_PI;
    double s12 = std::hypot(x2, y2);
    double lon2 = 0;
    double lat2 = 0;

    try {
        const auto& geod = GeographicLib::Geodesic::WGS84();
        geod.Direct(lat1, lon1, azi1, s12, lat2, lon2);
    }
    catch (const std::exception& e) {
        xll::log()->error("Caught exception: {}", e.what());
        result *= 0;
        return &result;
    }

    result(0, 0) = lon2;
    result(0, 1) = lat2;
    return &result;
}

/// Solves the inverse geodesic problem on the WGS84 ellipsoid.
/// \param[in] lon1 longitude at origin (degrees).
/// \param[in] lat1 latitude at origin (degrees).
/// \param[in] lon2 longitude at offset (degrees).
/// \param[in] lat2 latitude at offset (degrees).
/// \return array containing x-coordinate and y-coordinate at offset (meters).
XLL_EXPORT xll::static_fp12<2> * __stdcall geodesicInverse(
    const double lon1, const double lat1, const double lon2, const double lat2)
{
    thread_local xll::static_fp12<2> result(1, 2);
    
    double s12 = 0;
    double azi1 = 0;
    double azi2 = 0;
    
    try {
        const auto& geod = GeographicLib::Geodesic::WGS84();
        geod.Inverse(lat1, lon1, lat2, lon2, s12, azi1, azi2);
    }
    catch (const std::exception& e) {
        xll::log()->error("Caught exception: {}", e.what());
        result *= 0;
        return &result;
    }

    result(0, 0) = s12 * std::sin(azi1 * M_PI / 180.0);
    result(0, 1) = s12 * std::cos(azi1 * M_PI / 180.0);
    return &result;
}

/// Returns the GeographicLib version.
XLL_EXPORT const char * __stdcall libraryVersion(xll::variant *)
{
    const char *version = "GeographicLib " GEOGRAPHICLIB_VERSION_STRING;
    return version;
}

/// This function is called by the Add-in Manager to find the long name of the
/// add-in. If xAction = 1, this function should return a string containing the
/// long name of this XLL. If xAction = 2 or 3, this function should return
/// xlerrValue (#VALUE!).
XLL_EXPORT xll::variant * __stdcall xlAddInManagerInfo12(xll::variant *xAction)
{
    using namespace xll;
	
    thread_local variant xInfo, xIntAction;
    xloper<xlint> xDestType(xltypeInt);

	Excel12(xlCoerce, &xIntAction, xAction, &xDestType);
    
    if (xIntAction.xltype() == xltypeInt && xIntAction.get<xlint>().w == 1)
        xInfo.emplace<xlstr>(L"Geodesic Routines");
    else
        xInfo.emplace<xlerr>(error::excel_error::xlerrValue);
    
	return &xInfo;
}

/// Excel calls xlAutoOpen when the XLL is loaded. Register all functions and
/// perform any additional initialization in this function.
/// \return 1 on success, 0 on failure.
XLL_EXPORT int __stdcall xlAutoOpen()
{
    {
        xll::function_options opts;
        opts.argument_text = L"lon1,lat1,x2,y2";
        opts.category = L"Geodesic";
        opts.function_help = L"Solves the direct geodesic problem on the WGS84 ellipsoid.";
        opts.argument_help = {
            L"longitude at origin (degrees).",
            L"latitude at origin (degrees).",
            L"x-coordinate at offset (meters).",
            L"y-coordinate at offset (meters)."
        };
        xll::register_function(geodesicForward, L"geodesicForward", L"GEODESIC.FORWARD", opts);
    }

    {
        xll::function_options opts;
        opts.argument_text = L"lon1,lat1,lon2,lat2";
        opts.category = L"Geodesic";
        opts.function_help = L"Solves the inverse geodesic problem on the WGS84 ellipsoid.";
        opts.argument_help = {
            L"longitude at origin (degrees).",
            L"latitude at origin (degrees).",
            L"longitude at offset (degrees).",
            L"latitude at offset (degrees)."
        };
        xll::register_function(geodesicInverse, L"geodesicInverse", L"GEODESIC.INVERSE", opts);
    }

    {
        xll::function_options opts;
        opts.argument_text = L"arg";
        opts.category = L"Geodesic";
        opts.function_help = L"Returns the GeographicLib version.";
        opts.argument_help = { L"Argument ignored." };
        xll::register_function(libraryVersion, L"libraryVersion", L"GEODESIC.LIBVERSION", opts);
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
