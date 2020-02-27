//#define XLL_DISABLE_ASSERTS

#include <xll/xll.hpp>

#include <algorithm>
#include <array>
#include <iostream>
#include <string>
#include <vector>

#include <boost/core/lightweight_test.hpp>
#include <boost/numeric/ublas/io.hpp>

using namespace xll;

std::string xltype_name(XLTYPE xltype)
{
    switch (xltype) {
        case xltypeNum: return "xltypeNum";
        case xltypeStr: return "xltypeStr";
        case xltypeBool: return "xltypeBool";
        case xltypeRef: return "xltypeRef";
        case xltypeErr: return "xltypeErr";
        case xltypeFlow: return "xltypeFlow";
        case xltypeMulti: return "xltypeMulti";
        case xltypeMissing: return "xltypeMissing";
        case xltypeNil: return "xltypeNil";
        case xltypeSRef: return "xltypeSRef";
        case xltypeInt: return "xltypeInt";
        case xltypeBigData: return "xltypeBigData";
        default: return "xltypeUnknown";
    }
}

constexpr int test_constexpr()
{
    return variant12(value<tag::xlint>(555)).get<tag::xlint>().w;
}

void test_pstring()
{
    using namespace std::string_literals;

    auto fmtlit = make_pstring_literal("ABC");
    std::cout << fmtlit << "\n";

    auto comp1 = xloper12<tag::xlint>(888);
    auto comp2 = xloper12<tag::xlint>(889);
    std::cout << (comp1 == comp2) << "\n";

    std::cout << static_cast<std::string>(fmtlit) << "\n";
    std::cout << static_cast<std::string_view>(fmtlit) << "\n";

    static constexpr auto literal1 = make_wpstring_literal(L"Hello");
    std::wcout << literal1 << " " << literal1.size() << "\n";

    constexpr std::array<wchar_t, 5> arr {{ L'H', L'e', L'l', L'l', L'o' }};
    static constexpr auto literal2 = make_wpstring_array(arr);
    std::wcout << literal2 << " " << literal2.size() << "\n";

    std::cout << "*** pstring from std::string\n";
    pstring ps1(std::string("AAA"));
    std::cout << ps1 << " " << ps1.size() << "\n";
    
    std::cout << "*** pstring from std::wstring\n";
    pstring ps2(std::wstring(L"BBB"));
    std::cout << ps2 << " " << ps1.size() << "\n";
    
    std::cout << "*** wpstring from std::string\n";
    wpstring ps3(std::string("CCC"));
    std::cout << ps3 << " " << ps1.size() << "\n";
    
    std::cout << "*** wpstring from std::wstring\n";
    wpstring ps4(std::wstring(L"DDD"));
    std::cout << ps4 << " " << ps1.size() << "\n";

    fmt::print("{}\n", (std::string)ps1);

    std::cout << "*** value<tag::xlstr> test\n";
    auto strtest = value<tag::xlstr>(L"Test String"s);
    std::wcout << strtest << "\n";

    std::cout << "*** pstring test 1\n";
    
    std::cout << ps2 << " " << ps2.size() << "\n";
    std::wcout << ps3 << " " << ps3.size() << "\n";
    std::wcout << ps4 << " " << ps4.size() << "\n";

    auto s1 = static_cast<std::string>(ps1);
    auto s2 = static_cast<std::string>(ps2);
    auto s3 = static_cast<std::string>(ps3);
    auto s4 = static_cast<std::string>(ps4);
    auto ws1 = static_cast<std::wstring>(ps1);
    auto ws2 = static_cast<std::wstring>(ps2);
    auto ws3 = static_cast<std::wstring>(ps3);
    auto ws4 = static_cast<std::wstring>(ps4);

    std::cout << "*** pstring test 2\n";
    auto ps3c = std::move(ps3);
    auto ps3c2 = ps3;
    std::wcout << ps3c << " " << ps3c.size() << "\n";

    std::cout << "*** pstring test 3\n";
    std::cout << s1 << " " << s1.size() << "\n";
    std::cout << s2 << " " << s2.size() << "\n";
    std::cout << s3 << " " << s3.size() << "\n";
    std::cout << s4 << " " << s4.size() << "\n";
    std::wcout << ws1 << " " << ws1.size() << "\n";
    std::wcout << ws2 << " " << ws2.size() << "\n";
    std::wcout << ws3 << " " << ws3.size() << "\n";
    std::wcout << ws4 << " " << ws4.size() << "\n";

    std::cout << "*** pstring test 3\n";
    std::string_view ps1v(ps1);
    std::cout << ps1v << "\n";
    
}

void test_heap_detect()
{
    bool on_heap = false;

    {
        auto vv = value<tag::xlint>(555);
        variant12 *var1 = new variant12(vv);
        on_heap = var1->flags() & xlbitDLLFree;
        fmt::print("heap check 1: {}\n", on_heap);
    }
    {
        auto oper1 = xloper12<tag::xlint>(888);
        on_heap = oper1.flags() & xlbitDLLFree;
        fmt::print("heap check 2: {}\n", on_heap);
    }
    {
        auto oper2 = xloper12<tag::xlint>(value<tag::xlint>(444));
        on_heap = oper2.flags() & xlbitDLLFree;
        fmt::print("heap check 3: {}\n", on_heap);
    }
    {
        auto oper3 = new xloper12<tag::xlstr>(std::wstring(L"XXX"));
        on_heap = oper3->flags() & xlbitDLLFree;
        fmt::print("heap check 4: {}\n", on_heap);
    }
    {
        auto oper4 = new variant12();
        on_heap = oper4->flags() & xlbitDLLFree;
        fmt::print("heap check 5: {}\n", on_heap);
    }
    {
        fmt::print("Test Stack\n");
        variant12 test;
    }
    {
        fmt::print("Test Heap\n");
        auto test = std::make_unique<variant12>();
    }
}

int main()
{   
    namespace ublas = boost::numeric::ublas;
    namespace bc = boost::container;
    
    test_pstring();
    test_heap_detect();
}
