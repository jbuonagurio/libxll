//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <winnt.h>

#include <atomic>

// Global DLL support class, similar to CAtlModule (atlbase.h).
// COM objects call Lock() on construction and Unlock() on destruction,
// as they would when inheriting from CComObject (atlcom.h).

class XLLRibbonModule
{
public:
    static XLLRibbonModule& Instance()
    {
        static XLLRibbonModule instance;
        return instance;
    }
  
    HINSTANCE GetResourceInstance()
	{
		return m_hInstance;
	}

	void SetResourceInstance(HINSTANCE hInstance)
	{
		m_hInstance = hInstance;
	}
    
    void SetPerUserRegistration(bool perUser)
    {
        m_perUser = perUser;
    }
    
    bool GetPerUserRegistration()
    {
        return m_perUser;
    }
    
    unsigned int Lock()
	{
        return ++m_cLock;
	}

    unsigned int Unlock()
	{
        return --m_cLock;
	}

	unsigned int GetLockCount()
	{
		return m_cLock;
	}
  
private:
    XLLRibbonModule() = default;
    ~XLLRibbonModule() = default;
    XLLRibbonModule(const XLLRibbonModule&) = delete;
    XLLRibbonModule& operator=(const XLLRibbonModule&) = delete;

    HINSTANCE m_hInstance = nullptr;
    std::atomic_uint m_cLock{ 0 };
    bool m_perUser = true;
};

